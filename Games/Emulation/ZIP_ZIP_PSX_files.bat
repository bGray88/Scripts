for /D %%A in (*) do (
	cd %%A
	If Not Exist "%%A.zip" (
		for %%A in (*.img) do "C:\Program Files\7-Zip\7z.exe" a -tzip "%%~nA.zip" -xr!*.bat "%%A"
		del /f "%%A.img"
	)
	cd ..
)