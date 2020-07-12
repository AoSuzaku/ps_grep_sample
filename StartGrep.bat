echo off

REM PowerShellよりスクリプトを起動する。
REM 起動引数　第1引数：検索フォルダパス（絶対パス）　第2引数：検索ワード

powershell -executionpolicy remotesigned .\ExcelGrep.ps1 path 春

REM 実行結果確認のため、一時停止
PAUSE
