Dim objOutlookMsg,objOutlook
Set objOutlook = CreateObject("Outlook.Application")
Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
objOutlookMsg.To = ""
objOutlookMsg.CC = ""
objOutlookMsg.BCC = ""
objOutlookMsg.Subject = "test"
objOutlookMsg.Body = ""
objOutlookMsg.Attachments.DateAdd "?C:\Users\Tanbin\Desktop\Practice1.txt"
objOutlookMsg.Send

2) Dim FileLocation,myBrowser
myBrowser = "iexplore.exe"

myApp = "https://www.mlcalc.com/"'for file location variable
FileLocation = "C:\Users\Tanbin\Desktop\testSheet.xlsx"

DataTable.AddSheet "TestDataInUFT" 'add data sheet after local and global // runtime dataTable
DataTable.ImportSheet FileLocation,"TestDataForMC","TestDataInUFT"

TotalRows = DataTable.GetSheet("TestDataInUFT").GetRowCount()

For i = 1 To TotalRows
	DataTable.GetSheet("TestDataInUFT").SetCurrentRow(i) 'if i want to show value from changing row
	HomeValue = DataTable("pPrice","TestDataInUFT") 'works without,value as well
	DownPayment = DataTable("downPayment","TestDataInUFT")
	mortgage = DataTable("Mortgage","TestDataInUFT")
	interest_Rate = DataTable.Value("Interest_Rate","TestDataInUFT")
	property_TAX = DataTable("Property_TAX","TestDataInUFT")
	property_Insurance = DataTable("Property_TAX","TestDataInUFT")
	pMI = DataTable.Value("PMI","TestDataInUFT")
	zipCode = DataTable("ZipCode","TestDataInUFT")
	systemutil.Run myBrowser, myApp
	wait(3)
	Browser("title:=.*").Page("title:=.*").WebEdit("name:=ma").set HomeValue
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=dp").Set DownPayment
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mt").set mortgage
    Browser("title:=.*").page("title:=.*").WebEdit("name:=ir", "index:=0").Set interest_Rate
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pt").Set property_TAX
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pi").Set property_Insurance
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mi").Set pMI
    Browser("title:=.*").page("title:=.*").WebEdit("name:=zipCode","index:=0").Set zipCode
'    Browser("title:=.*").page("title:=.*").WebList("name:=sm", "index:=0").Select DataTable.Value("PaymentMonth","TestDataInUFT")
'    Browser("title:=.*").page("title:=.*").WebList("name:=sy","index:=0").Select DataTable.Value("PaymentYear","TestDataInUFT")
    wait(2)
    Browser("title:=.*").page("title:=.*").WebButton("type:=submit","index:=0").Click
    wait(3)
    Browser("title:=.*").page("title:=.*").WebElement("class:=big","html tag:=TD","innertest:=$1,591.04").highlight
'    Monthly_Payment = Browser("title:=.*").page("title:=.*").WebTable("name:=WebTable", "index:=3").GetCellData(1,1) 'for webtable
     Monthly_Payment = Browser("title:=.*").page("title:=.*").WebTable("name:=WebTable", "index:=3").GetCellData(1,1)
    
    datatable.Value("MonthlyPayment","TestDataInUFT") = Monthly_Payment
    
    Browser("title:=.*").Close
	
Next

DataTable.ExportSheet "C:\Users\Tanbin\Desktop\testResult.xlsx","TestDataInUFT"

'opening ie browser with app url

3) Dim FileLocation,myBrowser
myBrowser = "iexplore.exe"
myApp = "https://www.mlcalc.com/"'for file location variable

FileLocation = "C:\Users\Tanbin\Desktop\testSheet.xlsx"

DataTable.AddSheet "TestDataInUFT" 'add data sheet after local and global // runtime dataTable
DataTable.ImportSheet FileLocation,"TestDataForMC","TestDataInUFT"

TotalRows = DataTable.GetSheet("TestDataInUFT").GetRowCount()

For i = 1 To TotalRows
	DataTable.GetSheet("TestDataInUFT").SetCurrentRow(i) 'if i want to show value from changing row
	HomeValue = DataTable("pPrice","TestDataInUFT") 'works without,value as well
	DownPayment = DataTable("downPayment","TestDataInUFT")
	mortgage = DataTable("Mortgage","TestDataInUFT")
	interest_Rate = DataTable.Value("Interest_Rate","TestDataInUFT")
	property_TAX = DataTable("Property_TAX","TestDataInUFT")
	property_Insurance = DataTable("Property_TAX","TestDataInUFT")
	pMI = DataTable.Value("PMI","TestDataInUFT")
	zipCode = DataTable("ZipCode","TestDataInUFT")
	systemutil.Run myBrowser, myApp
	wait(3)
	Browser("title:=.*").Page("title:=.*").WebEdit("name:=ma").set HomeValue
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=dp").Set DownPayment
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mt").set mortgage
    Browser("title:=.*").page("title:=.*").WebEdit("name:=ir", "index:=0").Set interest_Rate
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pt").Set property_TAX
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pi").Set property_Insurance
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mi").Set pMI
    Browser("title:=.*").page("title:=.*").WebEdit("name:=zipCode","index:=0").Set zipCode
'    Browser("title:=.*").page("title:=.*").WebList("name:=sm", "index:=0").Select DataTable.Value("PaymentMonth","TestDataInUFT")
'    Browser("title:=.*").page("title:=.*").WebList("name:=sy","index:=0").Select DataTable.Value("PaymentYear","TestDataInUFT")
    wait(2)
    Browser("title:=.*").page("title:=.*").WebButton("type:=submit","index:=0").Click
    wait(3)
    Monthly_Payment = Browser("title:=.*").page("title:=.*").WebElement("class:=big","html tag:=TD","index:=0").GetTOProperties("innertext")
'    Monthly_Payment = Browser("title:=.*").page("title:=.*").WebTable("name:=WebTable", "index:=3").GetCellData(1,1) 'for webtable
'     Monthly_Payment = Browser("title:=.*").page("title:=.*").WebTable("name:=WebTable", "index:=3").GetCellData(1,1)
    
    datatable.Value("MonthlyPayment","TestDataInUFT") = Monthly_Payment
    
    Browser("title:=.*").Close
	
Next

DataTable.ExportSheet "C:\Users\Tanbin\Desktop\testResult.xlsx","TestDataInUFT"

'opening ie browser with app url
    
4) Dim FileLocation,myBrowser
myBrowser = "iexplore.exe"
myApp = "https://www.tdbank.com/"'for file location variable
SystemUtil.Run myBrowser,myApp

wait(2)
Browser("title:=.*").Page("title:=.*").Link("name:=Mortgage","index:=0").click
Browser("title:=.*").Page("title:=.*").WebList("name:=Q585").select "Purchase"
Browser("title:=.*").Page("title:=.*").WebEdit("name:=Q586").Set 650000
Browser("title:=.*").Page("title:=.*").WebEdit("name:=Q587").Set 670000
Browser("title:=.*").Page("title:=.*").WebList("name:=Q588").select "NY"
Browser("title:=.*").Page("title:=.*").WebList("html id:=Q589").select "Kings"
Browser("title:=.*").Page("title:=.*").WebButton("html id:=link-Get-Estimate").click
wait(3)
myMValue = Browser("title:=.*").Page("title:=.*").WebTable("name:=rate").WebElement("html tag:=SPAN","Index:=3").GetROProperty("innertext")
print myMValue
   
5) wBrowser = "iexplore.exe"
myApp = "http://demodt.cms.dealer.com/new-inventory/index.htm"

If Browser("title:=.*").Exist Then
Browser("title:=.*").Close
    
End If

systemutil.Run wBrowser, myApp
'working with checkbox
Browser("title:=.*").Page("title:=.*").WebCheckBox("name:=model","value:=Altima").Set "ON"
'working with radio button
Browser("title:=.*").Page("title:=.*").WebRadioGroup("name:=payment-selection").Select "payment-panel-paymentLease"
'working with image
Browser("title:=.*").Page("title:=.*").Image("href:=/new/Nissan/2018-Nissan-Altima-780581550a0e0ae84e674d4d4b942917.htm").Click

6) Dim myBrowser,myApp
myBrowser = "iexplore.exe"
myApp = "https://www.bestbuy.com/"

'If Browser("title:=.*").Exist Then
'Browser("title:=.*").Close   
'End If
'DataTable.AddSheet "TestDataInUFT"
systemutil.Run myBrowser, myApp
timestamp = Month(Now)&"-"&Year(Now)&"-"&Hour(Now)&"-"&Minute(Now)&"-"&Second(Now)
imageLocation = "C:\Users\Tanbin\Desktop\UFTScreenShots\GoogleHP"&timeStamp&".png"

Browser("title:=.*").page("title:=.*").Link("name:=Weekly Ad","index:=0").Click
Browser("title:=.*").page("title:=.*").Link("name:=Deal of the Day","index:=0").Click
wait(2)

'how to put a checkpoint to see the existance of an object

'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").Exist
'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").highlight
'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").GetROProperty("innertext")
'print myResult

if Browser("title:=.*").page("title:=.*").WebElement("html id:=resultStats", "index:=0").Exist then 
   reporter.ReportEvent micpass,"Total Search Result","Found It"
   Browser("title:=.*").page("title:=.*").CaptureBitmap imageLocation
   myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").GetROProperty(1,1)
   datatable.Value("SearchResult",dtglobalsheet) = myResult
   wait(2)
else
   reporter.ReportEvent micfail,"Total Search Result","Not Found"
   Browser("title:=.*").page("title:=.*").CaptureBitmap imageLocation
   myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").GetROProperty(1,1)
   datatable.Value("SearchResult",dtglobalsheet) = myResult
End if

'Browser("title:=.*").Close

7) Dim myBrowser,myApp
myBrowser = "iexplore.exe"
myApp = "https://www.google.com/"

'If Browser("title:=.*").Exist Then
'Browser("title:=.*").Close   
'End If
'DataTable.AddSheet "TestDataInUFT"
systemutil.Run myBrowser, myApp

Browser("title:=.*").page("title:=.*").webEdit("name:=q").Set "Automation Engineer"
Browser("title:=.*").page("title:=.*").webButton("name:=Google Search","index:=0").click
wait(2)

'how to put a checkpoint to see the existance of an object

'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").Exist
'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").highlight
'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").GetROProperty("innertext")
'print myResult

if Browser("title:=.*").page("title:=.*").WebElement("html id:=resultStats", "index:=0").Exist then 
   reporter.ReportEvent micpass,"Total Search Result","Found It"
   timestamp = Month(Now)&"-"&Year(Now)&"-"&Hour(Now)&"-"&Minute(Now)&"-"&Second(Now)
   imageLocation = "C:\Users\Tanbin\Desktop\UFTScreenShots\GoogleHP"&timeStamp&".png"
   Browser("title:=.*").page("title:=.*").CaptureBitmap imageLocation
   myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").GetROProperty(1,1)
   datatable.Value("SearchResult",dtglobalsheet) = myResult
   wait(2)
else
   reporter.ReportEvent micfail,"Total Search Result","Not Found"
   timestamp = Month(Now)&"-"&Year(Now)&"-"&Hour(Now)&"-"&Minute(Now)&"-"&Second(Now)
   imageLocation = "C:\Users\Tanbin\Desktop\UFTScreenShots\GoogleHP"&timeStamp&".png"
   Browser("title:=.*").page("title:=.*").CaptureBitmap imageLocation
End if

Browser("title:=.*").Close

8) Dim myBrowser,myApp
myBrowser = "iexplore.exe"
myApp = "https://www.google.com/"

'If Browser("title:=.*").Exist Then
'Browser("title:=.*").Close   
'End If
'DataTable.AddSheet "TestDataInUFT"
systemutil.Run myBrowser, myApp

Browser("title:=.*").page("title:=.*").webEdit("name:=q").Set "QA"
Browser("title:=.*").page("title:=.*").webButton("name:=Google Search","index:=0").click
wait(2)

'how to put a checkpoint to see the existance of an object

'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").Exist
'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").highlight
'myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").GetROProperty("innertext")
'print myResult

if Browser("title:=.*").page("title:=.*").WebElement("html id:=resultStats", "index:=0").Exist then 
   reporter.ReportEvent micpass,"Total Search Result","Found It"
   myResult = Browser("title:=.*").page("title:=.*").webElement("html id:=resultStats","index:=0").GetROProperty(1,1)
   datatable.Value("SearchResult",dtglobalsheet) = myResult
   wait(2)
else
   reporter.ReportEvent micfail,"Total Search Result","Not Found"
End if
wait(2)
'Browser("title:=.*").Close

9) myBrowser = "iexplore.exe"
myApp = "https://www.google.com"

systemutil.Run myBrowser, myApp

Call EnterValue("q", "Apple")

'Browser("title:=.*").Page("title:=.*").WebButton("name:=Google Search").Click
Call ClickButton("Google Search")



Sub ClickButton(ButtonName)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebButton"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("name") = ButtonName Then
            allObjects(i).click
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub



Sub EnterValue(EditName, inputValue)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebEdit"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("name") = EditName Then
            allObjects(i).set inputValue
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub

10) myBrowser = "iexplore.exe"
myApp = "https://www.mlcalc.com/"

systemutil.Run myBrowser, myApp

'Browser("title:=.*").Page("title:=.*").WebEdit("name:=ma").85000
'Browser("title:=.*").Page("title:=.*").WebEdit("name:=dp").25
'Browser("title:=.*").Page("title:=.*").WebEdit("name:=mt").set 40 
'Browser("title:=.*").page("title:=.*").WebEdit("name:=ir", "index:=0").Set 4.7   
'Browser("title:=.*").Page("title:=.*").WebEdit("name:=pt").Set 4500
'Browser("title:=.*").Page("title:=.*").WebEdit("name:=pi").Set 1800
'Browser("title:=.*").page("title:=.*").WebEdit("name:=zipCode","index:=0").Set 11218  

'Replacing above lines from 6-12 with below lines 15-21, but reusing same code. 
Call EnterValue("ma", 85000)
Call EnterValue("dp", 25)
Call EnterValue("mt", 40)
Call EnterValue("ir", 4.8)
Call EnterValue("pt", 4500)
Call EnterValue("pi", 1800)
Call EnterValue("zipCode", 11218)

'Set myPage = Browser("title:=.*").Page("title:=.*")
'Set myObject = Description.Create()
'    myObject("micClass").value = "WebEdit"
'    myObject("Name").value = "ma"
'
'myPage.WebEdit(myObject).Set 850000

Sub EnterValue(EditName, inputValue)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebEdit"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("name") = EditName Then
            allObjects(i).set inputValue
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub

11) number1 = Inputbox("enter a number")
number2 = Inputbox("enter a number")

totalResult = multipleValue(number1,number2)
print totalResult
finalResult = addValue(totalResult,10)
print finalResult
	
Function multipleValue(number1,number2)
	total1 = number1 * number2
	multipleValue = total1
End Function

Function addValue(a,b)
	total2 = a+b
	addValue = total2
End Function

12) BrowserName = "iexplore.exe"
AppName = "https://www.mlcalc.com/"

Call OpenApp(BrowserName, AppName)
call EnterValue("ma", 650000)
call EnterValue("dp", 20)
call EnterValue("mt", 20)
call EnterValue("ir", 4.7)
call EnterValue("pt", 3000)
call EnterValue("pi", 1500)
call EnterValue("hf", 2.4)
call EnterValue("mi", 0.7)
call EnterValue("zipCode", 11230)
Call EnterList("sm","Feb")
'Call EnterList("sy",2018)
call ClickButton("submit")
Call SnapShotsSite()
Call ClickLink("Mortgage Rates")
Call CloseApp


Function ClickLink(LinkName)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myLink = Description.Create()
        myLink("micClass").value = "Link"
    
    Set allLinks = myPage.ChildObjects(myLink)
        totalLinks = allLinks.count()
    
    For i = 0 To totalLinks-1
        'print allLinks(i).GetRoProperty("name")
        If allLinks(i).GetRoProperty("name") = LinkName Then
            allLinks(i).click
            Exit for 
        End If
        
        Set myPage = nothing
        Set myLink = nothing
        Set allLinks = nothing
        
    Next
    
End Function


Sub ClickButton(ButtonName)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebButton"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("type") = ButtonName Then
            allObjects(i).click
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub



Sub EnterValue(EditName, inputValue)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebEdit"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("name") = EditName Then
            allObjects(i).set inputValue
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub

Sub EnterList(EditList, inputValue)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebList"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("name") = EditList Then
            allObjects(i).select inputValue
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub

Function SnapShotsSite()
	timestamp = Month(Now)&"-"&Year(Now)&"-"&Hour(Now)&"-"&Minute(Now)&"-"&Second(Now)
    imageLocation = "C:\Users\Tanbin\Desktop\UFTScreenShots\GoogleHP"&timeStamp&".png"
    Browser("title:=.*").page("title:=.*").CaptureBitmap imageLocation
End Function

Sub OpenApp(myBrowser, myApp)

'    If Browser("title:=.*").Exist Then
'        Browser("title:=.*").Close
'    End If
    
    systemutil.Run myBrowser, myApp
    
End Sub


Sub CloseApp()
	Browser("title:=.*").Close
End Sub

13) BrowserName = "iexplore.exe"
AppName = "https://www.mlcalc.com/"


Call DataDrivenTesting()



Function ClickLink(LinkName)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myLink = Description.Create()
        myLink("micClass").value = "Link"
    
    Set allLinks = myPage.ChildObjects(myLink)
        totalLinks = allLinks.count()
    
    For i = 0 To totalLinks-1
        'print allLinks(i).GetRoProperty("name")
        If allLinks(i).GetRoProperty("name") = LinkName Then
            allLinks(i).click
            Exit for 
        End If
        
        Set myPage = nothing
        Set myLink = nothing
        Set allLinks = nothing
        
    Next
    
End Function


Sub ClickButton(ButtonName)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebButton"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("type") = ButtonName Then
            allObjects(i).click
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub



Sub EnterValue(EditName, inputValue)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebEdit"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("name") = EditName Then
            allObjects(i).set inputValue
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub

Sub EnterList(EditList, inputValue)
    
    Set myPage = Browser("title:=.*").Page("title:=.*")
    Set myObject = Description.Create()
        myObject("micClass").value = "WebList"
        'myObject("Name").value = EditName
    
    'myPage.WebEdit(myObject).Set inputValue
    
    Set allObjects = myPage.ChildObjects(myObject)
    totalObj = allObjects.count()
    
    For i = 0 To totalObj
        'print allObjects(i).GetROProperty("name")
        If allObjects(i).GetROProperty("name") = EditList Then
            allObjects(i).select inputValue
            Exit for 
        End If
        
    Next
    
    
    Set myPage = nothing
    Set myObject = nothing
    Set allObjects = nothing

    
End Sub

Function SnapShotsSite()
	timestamp = Month(Now)&"-"&Year(Now)&"-"&Hour(Now)&"-"&Minute(Now)&"-"&Second(Now)
    imageLocation = "C:\Users\Tanbin\Desktop\UFTScreenShots\GoogleHP"&timeStamp&".png"
    Browser("title:=.*").page("title:=.*").CaptureBitmap imageLocation
End Function

Function DataDrivenTesting()
	FileLocation = "C:\Users\Tanbin\Desktop\testSheet.xlsx"

	DataTable.AddSheet "TestDataInUFT" 
	DataTable.ImportSheet FileLocation,"TestDataForMC","TestDataInUFT"

	TotalRows = DataTable.GetSheet("TestDataInUFT").GetRowCount()

	For i = 1 To TotalRows
		DataTable.GetSheet("TestDataInUFT").SetCurrentRow(i) 
		Property_Price = DataTable("pPrice","TestDataInUFT") 
		Down_Payment = DataTable("downPayment","TestDataInUFT")
		Mortgage_Amount = DataTable("Mortgage","TestDataInUFT")
		Interest_Rate_Per = DataTable.Value("Interest_Rate","TestDataInUFT")
		Property_TAX_Per = DataTable("Property_TAX","TestDataInUFT")
		Property_Insurance_Per = DataTable("Property_Insurance","TestDataInUFT")
		PMI_Per = DataTable("PMI","TestDataInUFT")
		ZipCode = DataTable.Value("ZipCode","TestDataInUFT")
'		Select_Month = DataTable("sm","TestDataInUFT")
		Call OpenApp(BrowserName, AppName)
		wait(1)
		Browser("title:=.*").Page("title:=.*").WebEdit("name:=ma").set Property_Price
	    Browser("title:=.*").Page("title:=.*").WebEdit("name:=dp").Set Down_Payment
	    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mt").set Mortgage_Amount
	    Browser("title:=.*").page("title:=.*").WebEdit("name:=ir", "index:=0").Set Interest_Rate_Per
	    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pt").Set Property_TAX_Per
	    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pi").Set Property_Insurance_Per
	    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mi").Set PMI_Per
	    Browser("title:=.*").page("title:=.*").WebEdit("name:=zipCode","index:=0").Set ZipCode
'	    Browser("title:=.*").page("title:=.*").WebList("name:=sm","index:=0").Set Select_Month
	    
	    wait(2)
	    call ClickButton("submit")
	    Call SnapShotsSite()
	    wait(3)	    
	    Monthly_Payment = Browser("title:=.*").page("title:=.*").WebTable("name:=WebTable", "index:=3").GetCellData(1,1)	    
	    datatable.Value("MonthlyPayment","TestDataInUFT") = Monthly_Payment	   	    
	    call CloseApp
	  next
	    DataTable.ExportSheet "C:\Users\Tanbin\Desktop\testResult.xlsx","TestDataInUFT"
End Function

Sub OpenApp(myBrowser, myApp)

'    If Browser("title:=.*").Exist Then
'        Browser("title:=.*").Close
'    End If
    
    systemutil.Run myBrowser, myApp
    
End Sub


Sub CloseApp()
	Browser("title:=.*").Close
End Sub

14) bName = "iexplore.exe"
aName = "http://www.passport.gov.bd/"
Call OpenApp(bName, aName)
Call ClickOnCheck("ctl00$ContentPlaceHolder1$chkNext")
Call ClickButton("Continue to Online Enrolment")
Call SelectList("ctl00$ContentPlaceHolder1$ddCountry", "United States of America")
Call SelectList("ctl00$ContentPlaceHolder1$ddPassportType", "ORDINARY")
Call SelectRadioOption("ctl00$ContentPlaceHolder1$rdDeliveryType", "8")
Call SelectRadioOption("ctl00$ContentPlaceHolder1$rdGender","M")
EnterValue "ctl00$ContentPlaceHolder1$txtFullName", "Ashraful"
Call EnterValue("ctl00$ContentPlaceHolder1$txtFirstName", "Hoque")
EnterValue "ctl00$ContentPlaceHolder1$txtSurname", "Tanbin"
'call ClickOnCheck("ctl00$ContentPlaceHolder1$chkGuardian")
call EnterValue("ctl00$ContentPlaceHolder1$txtFatherName","Anisul")
Call SelectList("ctl00$ContentPlaceHolder1$ddCountry", "BANGLADESH")
Call SelectList("ctl00$ContentPlaceHolder1$ddCountry", "BANKER")
call EnterValue("ctl00$ContentPlaceHolder1$txtFatherName","Rahima")
Call SelectList("ctl00$ContentPlaceHolder1$ddCountry", "BANGLADESH")
Call SelectList("ctl00$ContentPlaceHolder1$ddCountry", "BANKER")

Function EnterValue(EditName, InputValue)

 Set myPage = Browser("title:=.*").Page("title:=.*")
 Set wEdit = Description.Create()
 wEdit("micClass").value = "WebEdit"

 Set allEdits = myPage.ChildObjects(wEdit)
 totalEdits = allEdits.count()
 print "total Edit Objects: "&totalEdits

 For i = 0 To totalEdits-1

 If allEdits(i).GetRoProperty("name") = EditName Then
 allEdits(i).Set InputValue
 wait(3)
 Exit for
 End If

 Next

 Set myPage = nothing
 Set wEdit = nothing
 Set totalEdits = nothing

End Function
Function SelectRadioOption(RadioName,RadioOption)

 Set myPage = Browser("title:=.*").Page("title:=.*")
 Set wRadio = Description.Create()
 wRadio("micClass").value = "WebRadioGroup"

 Set allRadios = myPage.ChildObjects(wRadio)
 totalRadios = allRadios.count()
 print "total Radio Objects: "&totalRadios

 For i = 0 To totalRadios-1

 If allRadios(i).GetRoProperty("name") = RadioName Then
 allRadios(i).select RadioOption
 wait(3)
 Exit for
 End If

 Next
 
 Set myPage = nothing
 Set wRadio = nothing
 Set allRadios = nothing

End Function
Function SelectList(ListName, ListOption)

 Set myPage = Browser("title:=.*").Page("title:=.*")
 Set wList = Description.Create()
 wList("micClass").value = "WebList"

 Set allLists = myPage.ChildObjects(wList)
 totalLists = allLists.count()
 print "total List Objects: "&totalLists

 For i = 0 To totalLists-1

 If allLists(i).GetRoProperty("name") = ListName Then
 allLists(i).select ListOption
 wait(3)
 Exit for
 End If

 Next

 Set myPage = nothing
 Set wList = nothing
 Set allLists = nothing

End Function
Function ClickButton(ButtonName)

 Set myPage = Browser("title:=.*").Page("title:=.*")
 Set wButton = Description.Create()
 wButton("micClass").value = "WebButton"

 Set allButtons = myPage.ChildObjects(wButton)
 totalButtons = allButtons.count()
 print "total webButton objects: "&totalButtons

 For i = 0 To totalButtons-1

 If allButtons(i).GetRoProperty("name") = ButtonName Then
 allButtons(i).click
 wait(5)
 Exit for
 End If

 Next

 Set myPage = nothing
 Set wButton = nothing
 Set allButtons = nothing
End Function
Function ClickOnCheck(CheckName)

 Set myPage = Browser("title:=.*").Page("title:=.*")
 Set webCheck = Description.Create()
 webCheck("micClass").value = "WebCheckBox"

 Set allChecks = myPage.ChildObjects(webCheck)
 totalChecks = allChecks.count()
 print "total webcheck Objects: "&totalChecks

 For i = 0 To totalChecks-1

 If allChecks(i).GetRoProperty("name") = CheckName Then
 allChecks(i).click
 Exit for
 End If

 Next

 Set myPage = nothing
 Set webCheck = nothing
 Set allChecks = nothing

End Function
Function OpenApp(BrowserName, appName) 'appName = http://www.passport.gov.bd/

' If Browser("title:=.*").Exist Then
' Browser("title:=.*").Close
' End If

 systemutil.Run BrowserName, appName

End Function


15) Function EnterValue(EditName, InputValue)

	 Set myPage = Browser("title:=.*").Page("title:=.*")
	 Set wEdit = Description.Create()
	 	wEdit("micClass").value = "WebEdit"
	
	 Set allEdits = myPage.ChildObjects(wEdit)
		 totalEdits = allEdits.count()
		 print "total Edit Objects: "&totalEdits
	
	 For i = 0 To totalEdits-1
	
	 If allEdits(i).GetRoProperty("name") = EditName Then
		 allEdits(i).Set InputValue
		 reporter.ReportEvent micPass,EditName,"worked"
		 wait(3)
	 Exit for
	 ElseIf i = totalEdits-1 Then
	 	 reporter.ReportEvent micFail,EditName,"Failed"
	 End If
	
	 Next
	
	 Set myPage = nothing
	 Set wEdit = nothing
	 Set totalEdits = nothing

End Function

Function SelectRadioOption(RadioName,RadioOption)

	 Set myPage = Browser("title:=.*").Page("title:=.*")
	 Set wRadio = Description.Create()
	 	wRadio("micClass").value = "WebRadioGroup"
	
	 Set allRadios = myPage.ChildObjects(wRadio)
		 totalRadios = allRadios.count()
		 print "total Radio Objects: "&totalRadios
	
	 For i = 0 To totalRadios-1
	
	 If allRadios(i).GetRoProperty("name") = RadioName Then
		 allRadios(i).select RadioOption
		 reporter.ReportEvent micPass,RadioName,"Worked"
		 wait(1)
	 Exit for
	 ElseIf i = totalEdits-1 Then
	 	reporter.ReportEvent micFail,RadioName,"Failed"
	 End If
	
	 Next
	 
	 Set myPage = nothing
	 Set wRadio = nothing
	 Set allRadios = nothing

End Function

Function SelectList(ListName, ListOption)

	 Set myPage = Browser("title:=.*").Page("title:=.*")
	 Set wList = Description.Create()
	 	wList("micClass").value = "WebList"
	
	 Set allLists = myPage.ChildObjects(wList)
		 totalLists = allLists.count()
		 print "total List Objects: "&totalLists
	
	 For i = 0 To totalLists-1
	
	 If allLists(i).GetRoProperty("name") = ListName Then
		 allLists(i).select ListOption
		 reporter.ReportEvent micPass,ListName,"Worked"
		 wait(1)
	 Exit for
	 ElseIf i = totalEdits-1 Then
	 	reporter.ReportEvent micFail,ListName,"Failed"
	 End If
	
	 Next
	
	 Set myPage = nothing
	 Set wList = nothing
	 Set allLists = nothing

End Function

Function ClickButtonName(ButtonName)

	 Set myPage = Browser("title:=.*").Page("title:=.*")
	 Set wButton = Description.Create()
	 	wButton("micClass").value = "WebButton"
	
	 Set allButtons = myPage.ChildObjects(wButton)
		 totalButtons = allButtons.count()
		 print "total webButton objects: "&totalButtons
	
	 For i = 0 To totalButtons-1
	
	 If allButtons(i).GetRoProperty("name") = ButtonName Then
		 allButtons(i).click
		 reporter.ReportEvent micPass,ButtonName,"Worked"
		 wait(1)
	 Exit for
	 ElseIf i = totalEdits-1 Then
	 	reporter.ReportEvent micFail,ButtonName,"Failed"
	 End If
	
	 Next
	
	 Set myPage = nothing
	 Set wButton = nothing
	 Set allButtons = nothing
	End Function
	
	
	Function ClickButtonType(ButtonName)

	 Set myPage = Browser("title:=.*").Page("title:=.*")
	 Set wButton = Description.Create()
	 	wButton("micClass").value = "WebButton"
	
	 Set allButtons = myPage.ChildObjects(wButton)
		 totalButtons = allButtons.count()
		 print "total webButton objects: "&totalButtons
	
	 For i = 0 To totalButtons-1
	
	 If allButtons(i).GetRoProperty("type") = ButtonName Then
		 allButtons(i).click
		 reporter.ReportEvent micPass,ButtonName,"Worked"
		 wait(1)
	 Exit for
	 ElseIf i = totalEdits-1 Then
	 	reporter.ReportEvent micFail,ButtonName,"Failed"
	 End If
	
	 Next
	
	 Set myPage = nothing
	 Set wButton = nothing
	 Set allButtons = nothing
	End Function
	
Function ClickOnCheck(CheckName)

	 Set myPage = Browser("title:=.*").Page("title:=.*")
	 Set webCheck = Description.Create()
	 	webCheck("micClass").value = "WebCheckBox"
	
	 Set allChecks = myPage.ChildObjects(webCheck)
		 totalChecks = allChecks.count()
		 print "total webcheck Objects: "&totalChecks
	
	 For i = 0 To totalChecks-1
	
	 If allChecks(i).GetRoProperty("name") = CheckName Then
		 allChecks(i).click
		 reporter.ReportEvent micPass,CheckName,"Worked"
	 Exit for
	 ElseIf i = totalEdits-1 Then
	 	reporter.ReportEvent micFail,CheckName,"Failed"
	 End If
	
	 Next
	
	 Set myPage = nothing
	 Set webCheck = nothing
	 Set allChecks = nothing

End Function

Function OpenApp(BrowserName, appName) 'appName = http://www.passport.gov.bd/

 systemutil.Run BrowserName, appName

End Function


Function SnapShotsSite()
	timestamp = Month(Now)&"-"&Year(Now)&"-"&Hour(Now)&"-"&Minute(Now)&"-"&Second(Now)
    imageLocation = "C:\Users\Tanbin\Desktop\UFTScreenShots\GoogleHP"&timeStamp&".png"
    Browser("title:=.*").page("title:=.*").CaptureBitmap imageLocation
End Function

Sub CloseApp()
	Browser("title:=.*").Close
End Sub

Function Monthly_Payment(FileLocation,MonthlyPayment)
	 Set myPage = Browser("title:=.*").Page("title:=.*")
	 Set myWebTable = Description.Create()
	 	myWebTable("micClass").value = "WebTable"
	
	 Set allWebTables = myPage.ChildObjects(myWebTable)
		 totalWebTables = allWebTables.count()
		 print "total webcheck Objects: "&totalWebTables
	
	 For i = 0 To totalWebTables-1
	 If allWebTables(i).GetRoProperty(1,1) = MonthlyPayment Then
		 datatable.Value("MonthlyPayment","TestDataInUFT") = Monthly_Payment
     Exit for  
	 End If
	 
	 next
	 Set myPage = nothing
	 Set myWebTable = nothing
	 Set allWebTables = nothing
End Function

Sub ShowingData(FileSaveLocation)
     DataTable.ExportSheet FileSaveLocation,"TestDataInUFT"	
End Sub



'main class


bName = "iexplore.exe"
aName = "https://www.mlcalc.com/"


	FileLocation = "C:\Users\Tanbin\Desktop\testSheet.xlsx"
    FileSaveLocation = "C:\Users\Tanbin\Desktop\testResult.xlsx"
	DataTable.AddSheet "TestDataInUFT" 
	DataTable.ImportSheet FileLocation,"TestDataForMC","TestDataInUFT"

	TotalRows = DataTable.GetSheet("TestDataInUFT").GetRowCount()

	For i = 1 To TotalRows
		DataTable.GetSheet("TestDataInUFT").SetCurrentRow(i) 
		Property_Price = DataTable("pPrice","TestDataInUFT") 
		Down_Payment = DataTable("downPayment","TestDataInUFT")
		Mortgage_Amount = DataTable("Mortgage","TestDataInUFT")
		Interest_Rate_Per = DataTable.Value("Interest_Rate","TestDataInUFT")
		Property_TAX_Per = DataTable("Property_TAX","TestDataInUFT")
		Property_Insurance_Per = DataTable("Property_Insurance","TestDataInUFT")
		PMI_Per = DataTable("PMI","TestDataInUFT")
		ZipCode = DataTable.Value("ZipCode","TestDataInUFT")
		PaymentMonth = DataTable.Value("PaymentMonth","TestDataInUFT")
		PaymentYear = DataTable.Value("PaymentYear","TestDataInUFT")
		call OpenApp(bName,aName)
		wait(2)
	    Call EnterValue("ma",Property_Price)
	    Call EnterValue("dp",Down_Payment)
	    Call EnterValue("mt",Mortgage_Amount)
	    Call EnterValue("ir",Interest_Rate_Per)
	    Call EnterValue("pt",Property_TAX_Per)
	    Call EnterValue("pi",Property_Insurance_Per)
	    Call EnterValue("mi",PMI_Per)
	    Call EnterValue("zipCode",ZipCode)
	    call SelectList("sm",PaymentMonth)
	    call SelectList("sy",PaymentYear)
	    call ClickButtonType("submit")
	    Call Monthly_Payment(FileLocation,MonthlyPayment)
	    Call ShowingData(FileSaveLocation)
	    Call CloseApp()
'	    Call SelectList("sm")
	next 	    
'	    Monthly_Payment = Browser("title:=.*").page("title:=.*").WebTable("name:=WebTable", "index:=3").GetCellData(1,1)	    
'	    datatable.Value("MonthlyPayment","TestDataInUFT") = Monthly_Payment	   	    
'	    
'	    next
'	    DataTable.ExportSheet "C:\Users\Tanbin\Desktop\testResult.xlsx","TestDataInUFT"

16) FileLocation ="C:\Users\Tanbin\Desktop\Bing_Hybrid_Framework.xls"

DataTable.AddSheet "TestDataInUFT"' to create a datatable during runtime datatable 

DataTable.ImportSheet FileLocation,"BingTest","TestDataInUFT"

TotalTestSteps = DataTable.GetSheet("TestDataInUFT").GetRowCount()



For i = 1  To TotalTestSteps
    
        DataTable.GetSheet("TestDataInUFT").SetCurrentRow(i)
    
        KeyValue = datatable.Value("Keyword","TestDataInUFT")
        ObjValue = datatable.Value("Obj_info","TestDataInUFT")
        InputValue = datatable.Value("Obj_Input","TestDataInUFT")
        
        Call ActionTaker(KeyValue,ObjValue,InputValue)
    
'    Select Case KeyWord
'    
'        Case "OpenApp"
'        	call OpenApp(ObjInfo,InputInfo)
'        	
'    	Case "EnterValue"
'    		call EnterValue(ObjInfo,InputInfo)
'    		
'    	Case "ClickButton"
'    		call ClickButtonName(ObjInfo)
'    		
'    	Case "ClickLink"
'    		Call ClickLink(ObjInfo)
'    		
'    	Case "CloseApp"
'    		Call CloseApp()
'    	
'    End Select
    
'    If KeyWord = "OpenApp" Then
'        call OpenApp(ObjInfo,InputInfo)
'        
'    End if 
'        
'    If KeyWord = "EnterValue" Then
'        call EnterValue(ObjInfo,InputInfo)
'            
'    End If
'
'    If KeyWord = "ClickButton" Then
'        call ClickButton(ObjInfo)
'    End If     
'    
'    If KeyWord = "CloseApp" Then
'        Call CloseApp()
'        
'    End if
'    
'    If KeyWord = "ClickLink" Then
'        Call ClickLink(ObjInfo)
'                    
'    End If
    
Next


'PS_VBFUnction



Function OpenApp(BrowserName,appName) 

   If Browser("title:=.*").Exist Then
         Browser("title:=.*").close
   End If

    systemutil.Run BrowserName, appName
    
End Function

Sub CloseApp()
    
    Browser("title:=.*").Close
End Sub



Function EnterValue(EditName,inputValue)
    
    Set myPage= Browser("title:=.*").page("title:=.*")
    Set wEdit = description.Create() 
        wEdit("micClass").value = "WebEdit"
        
    Set allEdits = myPage.ChildObjects(wEdit) 
        totalEdits = allEdits.count()
    print totalEdits
     
     
     For i = 0 To totalEdits-1
     
        If allEdits(i).GetROProperty("name") = EditName Then
            allEdits(i).set inputValue
            reporter.ReportEvent micPass, EditName , "Worked"
            wait(3)
            Exit for
            
        ElseIf i = totalEdits-1 Then
            reporter.ReportEvent micFail, EditName , "Failed"
            
        End If
         
     Next
    
    
     Set myPage= nothing
     Set wEdit = nothing
     Set totalEdits = nothing
     
     
End Function


Function SelectRadioOption(RadioName,RadioOption)
    
    Set myPage= Browser("title:=.*").page("title:=.*")
    Set wRadio = description.Create() 
        wRadio("micClass").value = "WebRadioGroup"
        
        
    Set allRadios = myPage.ChildObjects(wRadio) 
        totalRadios = allRadios.count()
    print totalRadios
     
     
     For i = 0 To totalRadios-1
     
        If allRadios(i).GetROProperty("name") = RadioName Then
            allRadios(i).select RadioOption
            wait(3)
            reporter.ReportEvent micPass, RadioName , "Worked"
            Exit for
            
        ElseIf i = totalRadios-1 Then
            reporter.ReportEvent micFail, RadioName , "Failed"
        End If
         
     Next
    
    
     Set myPage= nothing
     Set wRadio = nothing
     Set totalRadios = nothing
     
     
End Function


Function SelectList(ListName,ListOption)
    
    Set myPage= Browser("title:=.*").page("title:=.*")
    Set wLists = description.Create() 
        wLists("micClass").value = "WebList"
        
    Set allLists = myPage.ChildObjects(wLists) 
        totalLists = allLists.count()
    print totalLists
     
     
     For i = 0 To totalLists-1
     
        If allLists(i).GetROProperty("name") = ListName Then
            allLists(i).select ListOption
            reporter.ReportEvent micPass, ListName , "Worked"
            wait(3)
            
            Exit for
            
         ElseIf i = totalLists-1 Then
            reporter.ReportEvent micFail, ListName , "Failed"    
            
        End If
         
     Next
    
    
     Set myPage= nothing
     Set wLists = nothing
     Set totalLists = nothing
     
     
End Function


Function ClickLink(LinkName)
    
    Set myPage= Browser("title:=.*").page("title:=.*")
    Set wELink = description.Create() 
        wELink("micClass").value = "Link"
        
    Set allLinks = myPage.ChildObjects(wELink) 
        totalLionks = allLinks.count()
    print totalLionks
     
     
     For i = 0 To totalLionks-1
     
        If allLinks(i).GetROProperty("name") = LinkName Then
            allLinks(i).click
        reporter.ReportEvent micPass, LinkName , "Worked"
            Exit for
            
       ElseIf i= totalLionks-1 Then 
             reporter.ReportEvent micFail, LinkName                                                                                                                                                                                                                                 , "Failed"
             
             
        End If
         
     Next
    
    
     Set myPage= nothing
     Set wButton = nothing
     Set totalLionks = nothing
     
     
End Function



Function ClickButton2(ObjInfo)
    
    Set myPage= Browser("title:=.*").page("title:=.*")
    Set wButton2 = description.Create() 
        wButton2("micClass").value = "WebButton"
        
    Set allButtons2 = myPage.ChildObjects(wButton2) 
        totalButtons2 = allButtons2.count()
    print totalButtons
     
     
     For i = 0 To totalButtons2-1
     
        If allButtons2(i).GetROProperty("type") = ObjInfo Then
            allButtons2(i).click
        reporter.ReportEvent micPass, ObjInfo , "Worked"
            Exit for
            
       ElseIf i= totalButtons-1 Then 
             reporter.ReportEvent micFail, ObjInfo                                                                                                                                                                                                                                 , "Failed"
             
             
        End If
         
     Next
    
    
     Set myPage= nothing
     Set wELink = nothing
     Set totalButtons = nothing
     
     
End Function


Function ClickButton(ButtonName)
    
    Set myPage= Browser("title:=.*").page("title:=.*")
    Set wButton = description.Create() 
        wButton("micClass").value = "WebButton"
        
    Set allButtons = myPage.ChildObjects(wButton) 
        totalButtons = allButtons.count()
    print totalButtons
     
     
     For i = 0 To totalButtons-1
     
        If allButtons(i).GetROProperty("name") = ButtonName Then
            allButtons(i).click
        reporter.ReportEvent micPass, ButtonName , "Worked"
            Exit for
            
       ElseIf i= totalButtons-1 Then 
             reporter.ReportEvent micFail, ButtonName , "Failed"
             
             
        End If
         
     Next
    
    
     Set myPage= nothing
     Set wButton = nothing
     Set totalButtons = nothing
     
     
End Function


Function ClickonCheck(CheckName)
    
    Set myPage= Browser("title:=.*").page("title:=.*")
    Set webCheck = description.Create() 
        webCheck("micClass").value = "WebCheckBox"
        
    Set allChecks = myPage.ChildObjects(webCheck) 
        totalChecks = allChecks.count()
    print totalChecks
     
     
     For i = 0 To totalChecks-1
     
        If allChecks(i).GetROProperty("name") = CheckName Then
            allChecks(i).click
            reporter.ReportEvent micPass, CheckName , "Worked"
            Exit for
            
          ElseIf i= totalChecks-1 Then 
             reporter.ReportEvent micFail, CheckName , "Failed"
        End If
         
     Next
    
    
     Set myPage= nothing
     Set webCheck = nothing
     Set allChecks = nothing
     
     
End Function


'UtilityScript


Function ActionTaker(KeyWord,ObjInfo,InputInfo)
	
	Select Case KeyWord
    
        Case "OpenApp"
        	call OpenApp(ObjInfo,InputInfo)
        	
    	Case "EnterValue"
    		call EnterValue(ObjInfo,InputInfo)
    		
    	Case "ClickButton"
    		call ClickButton(ObjInfo)
    		
    	Case "ClickLink"
    		Call ClickLink(ObjInfo)
    		
    	Case "CloseApp"
    		Call CloseApp()
    	
  	End Select
  
End Function



