# Rollbar-Classic-ASP
Rollbar.com - Classic ASP API

A simpleVBScript / Classic ASP wrapper for the www.rollbar.com web API. 

How to use: 
1. Fill our the strRollbarAccessToken in rollbar.asp
1. On your custom 500 error page, include rollbar.asp and call `RollbarASPError()`

To log items in rollbar manually use `RollbarError(strMessage, strExtraPayload)` / `RollbarWarning(strMessage, strExtraPayload)` / `RollbarInfo(strMessage, strExtraPayload)` / `RollbarDebug(strMessage, strExtraPayload)` accordingly.