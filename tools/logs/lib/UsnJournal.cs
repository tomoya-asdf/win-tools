using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace WinTools
{
    public sealed class FolderChange
    {
        public DateTime TimeCreated { get; set; }
        public string Operation { get; set; }
        public string Name { get; set; }
        public string ParentPathNow { get; set; }
        public string PathHint { get; set; }
        public ulong FileId { get; set; }
        public ulong ParentFileId { get; set; }
        public long Usn { get; set; }
        public uint ReasonMask { get; set; }
    }

    // Read-only NTFS journal access. No journal creation, deletion or audit changes.
    public static class UsnJournal
    {
        private const uint ChangeMask = 0x0010BB00;
        private const uint QueryJournal = 0x000900F4;
        private const uint ReadJournal = 0x000900BB;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string name, uint access,
            uint share, IntPtr security, uint creation, uint flags, IntPtr template);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool DeviceIoControl(SafeFileHandle handle, uint code,
            byte[] input, int inputSize, [Out] byte[] output, int outputSize,
            out int bytesReturned, IntPtr overlapped);

        // FILE_ID_DESCRIPTOR includes a 16-byte union, including on 32-bit Windows.
        [StructLayout(LayoutKind.Explicit, Size = 24)]
        private struct FileIdDescriptor
        {
            [FieldOffset(0)] public uint Size;
            [FieldOffset(4)] public uint Type;
            [FieldOffset(8)] public ulong Id;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern SafeFileHandle OpenFileById(SafeFileHandle volume,
            ref FileIdDescriptor id, uint access, uint share, IntPtr security, uint flags);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFinalPathNameByHandle(SafeFileHandle handle,
            StringBuilder path, uint size, uint flags);

        private static SafeFileHandle OpenVolume(string drive)
        {
            SafeFileHandle handle = CreateFile(@"\\.\" + drive, 0x80000000,
                7, IntPtr.Zero, 3, 0, IntPtr.Zero);
            if (handle.IsInvalid)
            {
                int error = Marshal.GetLastWin32Error();
                handle.Dispose();
                throw new Win32Exception(error, "Cannot open volume " + drive);
            }
            return handle;
        }

        private static string ParentPath(SafeFileHandle volume, ulong parentId)
        {
            FileIdDescriptor id = new FileIdDescriptor { Size = 24, Type = 0, Id = parentId };
            // Zero access queries metadata; BACKUP_SEMANTICS permits directories.
            using (SafeFileHandle parent = OpenFileById(volume, ref id, 0, 7,
                IntPtr.Zero, 0x02000000))
            {
                if (parent.IsInvalid) return null;
                StringBuilder path = new StringBuilder(32768);
                uint length = GetFinalPathNameByHandle(parent, path, (uint)path.Capacity, 0);
                if (length == 0 || length >= path.Capacity) return null;
                string result = path.ToString();
                return result.StartsWith(@"\\?\", StringComparison.Ordinal)
                    ? result.Substring(4) : result;
            }
        }

        private static string Describe(uint reason)
        {
            List<string> names = new List<string>();
            if ((reason & 0x100) != 0) names.Add("Created");
            if ((reason & 0x200) != 0) names.Add("Deleted");
            if ((reason & 0x1000) != 0) names.Add("RenamedFrom");
            if ((reason & 0x2000) != 0) names.Add("RenamedTo");
            if ((reason & 0x800) != 0) names.Add("SecurityChanged");
            if ((reason & 0x8000) != 0) names.Add("MetadataChanged");
            if ((reason & 0x100000) != 0) names.Add("ReparsePointChanged");
            return String.Join(", ", names.ToArray());
        }

        // Separately callable to verify binary parsing without volume privileges.
        // NTFS V2 is requested explicitly; unknown versions must never be misparsed.
        public static List<FolderChange> ParseBuffer(byte[] buffer, int count, long endUsn)
        {
            if (buffer == null || count < 8 || count > buffer.Length)
                throw new InvalidDataException("Invalid USN buffer size.");
            List<FolderChange> result = new List<FolderChange>();
            int offset = 8; // First 8 bytes contain the continuation USN.
            while (offset < count)
            {
                if (count - offset < 8) throw new InvalidDataException("Truncated USN header.");
                uint size = BitConverter.ToUInt32(buffer, offset);
                ushort version = BitConverter.ToUInt16(buffer, offset + 4);
                if (size < 60 || size > count - offset)
                    throw new InvalidDataException("Invalid USN record length.");
                if (version != 2)
                    throw new NotSupportedException("Unsupported USN record version: " + version);
                ushort nameLength = BitConverter.ToUInt16(buffer, offset + 56);
                ushort nameOffset = BitConverter.ToUInt16(buffer, offset + 58);
                if (nameOffset < 60 || (nameLength % 2) != 0 || nameOffset + nameLength > size)
                    throw new InvalidDataException("Invalid USN file name bounds.");
                long usn = BitConverter.ToInt64(buffer, offset + 24);
                uint reason = BitConverter.ToUInt32(buffer, offset + 40);
                uint attributes = BitConverter.ToUInt32(buffer, offset + 52);
                if (usn < endUsn && (attributes & 0x10) != 0 && (reason & ChangeMask) != 0)
                {
                    result.Add(new FolderChange {
                        FileId = BitConverter.ToUInt64(buffer, offset + 8),
                        ParentFileId = BitConverter.ToUInt64(buffer, offset + 16),
                        Usn = usn,
                        TimeCreated = DateTime.FromFileTimeUtc(BitConverter.ToInt64(buffer, offset + 32)).ToLocalTime(),
                        ReasonMask = reason,
                        Operation = Describe(reason),
                        Name = Encoding.Unicode.GetString(buffer, offset + nameOffset, nameLength)
                    });
                }
                offset += (int)size;
            }
            return result;
        }

        public static FolderChange[] Read(string drive, DateTime start, DateTime end,
            int maxEvents, string nameContains)
        {
            if (drive == null || drive.Length != 2 || !Char.IsLetter(drive[0]) || drive[1] != ':')
                throw new ArgumentException("Use a drive letter such as C:.");
            if (maxEvents < 1 || start > end) throw new ArgumentException("Invalid range or limit.");
            using (SafeFileHandle volume = OpenVolume(drive))
            {
                byte[] journal = new byte[96];
                int returned;
                if (!DeviceIoControl(volume, QueryJournal, null, 0, journal,
                    journal.Length, out returned, IntPtr.Zero))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot query existing USN journal.");
                if (returned < 56) throw new InvalidDataException("Truncated journal information.");
                ulong journalId = BitConverter.ToUInt64(journal, 0);
                long cursor = Math.Max(BitConverter.ToInt64(journal, 8), BitConverter.ToInt64(journal, 24));
                long snapshotEnd = BitConverter.ToInt64(journal, 16);
                Queue<FolderChange> latest = new Queue<FolderChange>();
                byte[] buffer = new byte[1024 * 1024];
                // READ_USN_JOURNAL_DATA_V1: V2 records only, no wait at the tail.
                byte[] input = new byte[48];
                Array.Copy(BitConverter.GetBytes(ChangeMask), 0, input, 8, 4);
                Array.Copy(BitConverter.GetBytes(journalId), 0, input, 32, 8);
                Array.Copy(BitConverter.GetBytes((ushort)2), 0, input, 40, 2);
                Array.Copy(BitConverter.GetBytes((ushort)2), 0, input, 42, 2);
                // Read up to the initial tail, even if other processes keep writing.
                while (cursor < snapshotEnd)
                {
                    Array.Copy(BitConverter.GetBytes(cursor), 0, input, 0, 8);
                    if (!DeviceIoControl(volume, ReadJournal, input, input.Length,
                        buffer, buffer.Length, out returned, IntPtr.Zero))
                        throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot read USN journal; it may have rotated. Retry the command.");
                    foreach (FolderChange change in ParseBuffer(buffer, returned, snapshotEnd))
                    {
                        if (change.TimeCreated < start || change.TimeCreated > end) continue;
                        if (!String.IsNullOrEmpty(nameContains) &&
                            change.Name.IndexOf(nameContains, StringComparison.OrdinalIgnoreCase) < 0) continue;
                        latest.Enqueue(change);
                        if (latest.Count > maxEvents) latest.Dequeue();
                    }
                    long next = BitConverter.ToInt64(buffer, 0);
                    if (next <= cursor) break;
                    cursor = next;
                }
                FolderChange[] result = latest.ToArray();
                Array.Reverse(result);
                Dictionary<ulong, string> parents = new Dictionary<ulong, string>();
                foreach (FolderChange change in result)
                {
                    string parent;
                    if (!parents.TryGetValue(change.ParentFileId, out parent))
                    {
                        parent = ParentPath(volume, change.ParentFileId);
                        parents.Add(change.ParentFileId, parent);
                    }
                    change.ParentPathNow = parent;
                    // This is deliberately a hint: ancestors may have been renamed.
                    change.PathHint = parent == null ? null : Path.Combine(parent, change.Name);
                }
                return result;
            }
        }
    }
}
