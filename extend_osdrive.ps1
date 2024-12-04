# Iterate through all the disks on the Windows machine
foreach($disk in Get-Disk)
{
    # Check if the disk in context is a Boot and System disk
    if((Get-Disk -Number $disk.number).IsBoot -And (Get-Disk -Number $disk.number).IsSystem)
    {
        # Get the drive letter assigned to the disk partition where OS is installed
        $driveLetter = (Get-Partition -DiskNumber $disk.Number | where {$_.DriveLetter}).DriveLetter
        "Current OS Drive: $driveLetter :\"

        # Get current size of the OS parition on the Disk
        $currentOSDiskSize = (Get-Partition -DriveLetter $driveLetter).Size        
        "Current OS Partition Size: $currentOSDiskSize"

        # Get Partition Number of the OS partition on the Disk
        $partitionNum = (Get-Partition -DriveLetter $driveLetter).PartitionNumber
        "Current OS Partition Number: $partitionNum"

        # Get the available unallocated disk space size
        $unallocatedDiskSize = (Get-Disk -Number $disk.number).LargestFreeExtent
        "Total Unallocated Space Available: $unallocatedDiskSize"

        # Get the max allowed size for the OS Partition on the disk
        $allowedSize = (Get-PartitionSupportedSize -DiskNumber $disk.Number -PartitionNumber $partitionNum).SizeMax
        "Total Partition Size allowed: $allowedSize"

        if ($unallocatedDiskSize -gt 0 -And $unallocatedDiskSize -le $allowedSize)
        {
            $totalDiskSize = $allowedSize
            
            # Resize the OS Partition to Include the entire Unallocated disk space
            $resizeOp = Resize-Partition -DriveLetter C -Size $totalDiskSize
            "OS Drive Resize Completed $resizeOp"
        }
        else {
            "There is no Unallocated space to extend OS Drive Partition size"
        }
    }   
}