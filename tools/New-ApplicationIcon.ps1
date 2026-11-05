param(
    [string]$Png=(Join-Path $PSScriptRoot '../ui/assets/email-archive.png'),
    [string]$Output=(Join-Path $PSScriptRoot '../ui/assets/email-archive.ico')
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
$source=[Drawing.Bitmap]::new([IO.Path]::GetFullPath($Png))
$images=[Collections.Generic.List[byte[]]]::new()
$sizes=@(16,24,32,48,64,128,256)
try{
    foreach($size in $sizes){
        $bitmap=[Drawing.Bitmap]::new($size,$size,[Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics=[Drawing.Graphics]::FromImage($bitmap)
        $memory=[IO.MemoryStream]::new()
        try{
            $graphics.CompositingMode=[Drawing.Drawing2D.CompositingMode]::SourceCopy
            $graphics.InterpolationMode=[Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode=[Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.Clear([Drawing.Color]::Transparent)
            $graphics.DrawImage($source,[Drawing.Rectangle]::new(0,0,$size,$size))
            $bitmap.Save($memory,[Drawing.Imaging.ImageFormat]::Png)
            $images.Add($memory.ToArray())
        }finally{$memory.Dispose();$graphics.Dispose();$bitmap.Dispose()}
    }
    $stream=[IO.File]::Create([IO.Path]::GetFullPath($Output));$writer=[IO.BinaryWriter]::new($stream)
    try{
        $writer.Write([uint16]0);$writer.Write([uint16]1);$writer.Write([uint16]$sizes.Count)
        $offset=6+16*$sizes.Count
        for($i=0;$i -lt $sizes.Count;$i++){
            $dimension=if($sizes[$i] -eq 256){0}else{$sizes[$i]}
            $writer.Write([byte]$dimension);$writer.Write([byte]$dimension);$writer.Write([byte]0);$writer.Write([byte]0)
            $writer.Write([uint16]1);$writer.Write([uint16]32);$writer.Write([uint32]$images[$i].Length);$writer.Write([uint32]$offset)
            $offset+=$images[$i].Length
        }
        foreach($bytes in $images){$writer.Write([byte[]]$bytes)}
    }finally{$writer.Dispose();$stream.Dispose()}
}finally{$source.Dispose()}
Write-Output ('Ícone multirresolução criado em '+[IO.Path]::GetFullPath($Output))
