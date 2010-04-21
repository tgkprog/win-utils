rem from 0 to 100
set volumeNow=0
rem seconds
set waitSeconds=11
rem from 0 to 100
set volumeLater=100
start VolumeSetOnTimerChange.exe %volumeNow% %waitSeconds% %volumeLater%