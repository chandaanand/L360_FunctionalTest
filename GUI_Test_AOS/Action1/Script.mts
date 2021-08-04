'This Action module used to launch AOS using Chrome and Login
SystemUtil.Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
'SystemUtil.Run "C:\Program Files\internet explorer\iexplore.exe"
Browser("Browser").Navigate "https://advantageonlineshopping.com/#/"
Browser("Browser").Page("Advantage Shopping").Link("Profile_Link").Click @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("Browser").Page("Advantage Shopping").WebEdit("username").Set DataTable("aos_uname", dtLocalSheet) @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Browser").Page("Advantage Shopping").WebEdit("password").SetSecure DataTable("aos_pwd", dtLocalSheet) @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Browser").Page("Advantage Shopping").WebButton("Btn_sign_in").Click @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Browser").Page("Advantage Shopping").WebElement("Loggedin_UserName").Check CheckPoint("admin") @@ script infofile_;_ZIP::ssf5.xml_;_

