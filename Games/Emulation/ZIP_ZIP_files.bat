for %%A in (*) do "C:\Program Files\7-Zip\7z.exe" a -tzip "%%~nA.zip" -xr!*.bat "%%A"