for %%A in (*.iso) do "C:\Program Files\7-Zip\7z.exe" a -t7z -m0=lzma2:d27 -mx9 -mmt8 "%%~nA.7z" -xr!*.bat "%%A"