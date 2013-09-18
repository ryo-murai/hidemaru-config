hidemaru-config
================
秀丸エディタのポータブルバージョンで使用される設定ファイル(`HmRegIni.ini`)やその他(`hilight`)ファイル等を置いておく場所

ポータブルバージョンの秀丸エディタは、_秀丸エディタ持ち出しキット_ で作成する必要がある

```
git clone https://github.com/ryo-murai/hidemaru-config.git

cd hidemaru-config

@powershell -NoProfile -ExecutionPolicy unrestricted -Command "iex ((new-object net.webclient).DownloadString('https://raw.github.com/ryo-murai/ChocolateyPackages/master/hidemaru-portable/tools/InstallHidemaruPortable.ps1'))"

```
