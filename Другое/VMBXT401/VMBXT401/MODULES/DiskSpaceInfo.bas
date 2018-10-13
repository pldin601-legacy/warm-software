Attribute VB_Name = "DiskSpaceInfo"
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long

Function GotFreeDiskSpace(DriveName As String)
 Dim RetVal, SectPerCls, BtPerSect, FreeCls, TotalCls
 A = GetDiskFreeSpace(DriveName, SectPerCls, BtPerSect, FreeCls, TotalCls)
 GotFreeDiskSpace = SectPerCls * BtPerSect * FreeCls
End Function

Function GotTotalDiskSpace(DriveName As String)
 Dim RetVal, SectPerCls, BtPerSect, FreeCls, TotalCls
 A = GetDiskFreeSpace(DriveName, SectPerCls, BtPerSect, FreeCls, TotalCls)
 GotTotalDiskSpace = SectPerCls * BtPerSect * TotalCls
End Function

