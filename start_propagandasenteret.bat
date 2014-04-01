@ECHO off
SETLOCAL EnableDelayedExpansion 

:: Create a temporary drive letter mapped to your UNC root location
:: and effectively CD to that location

SET cur_path=%~dp0
echo %cur_path%
PUSHD %cur_path%

:: Do some work

SET app_path=%appdata%\Scriptotek\Propagandasenteret
SET mda=0
SET mdb=0

:: Force copy

IF NOT EXIST %appdata%\Scriptotek\NUL (
	MKDIR %appdata%\Scriptotek
)
IF NOT EXIST %app_path%\NUL (
	MKDIR %app_path%
)

IF NOT EXIST %app_path%\propagandasenteret.hta (
	ECHO First run. Copying files
	COPY propagandasenteret.hta %app_path%\propagandasenteret.hta
	COPY jquery-1.10.1.min.js %app_path%\jquery-1.10.1.min.js
	COPY Broadcast.ico %app_path%\icon.ico
) ELSE (
	FOR /f "skip=3" %%G IN ('fciv propagandasenteret.hta') DO (
		SET mda=%%G
	)
	
	FOR /f "skip=3" %%b IN ('fciv %app_path%\propagandasenteret.hta') DO (
		SET mdb=%%b
	)
	
	ECHO !mda! !mdb!

	IF !mda!==!mdb! (
		ECHO File unchanged. No update needed
	) ELSE (
		ECHO File changed. Updating
		COPY propagandasenteret.hta %app_path%\propagandasenteret.hta
	)

)

:: Remove the temporary drive letter and return to your original location
POPD

START %app_path%\propagandasenteret.hta
