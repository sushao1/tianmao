on run
    set appFolder to POSIX path of (path to me)
    set toolPath to appFolder & "Contents/Resources/图片批量压缩工具.html"
    do shell script "/usr/bin/open " & quoted form of toolPath
end run
