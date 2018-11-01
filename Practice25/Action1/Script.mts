'Global variable(declare global variable to envirom=nment variable)
Environment("StepStatus") = ""
Environment("TotalStep_Status") = 0
Environment("TotalTC") = 0

FileLocation ="C:\Users\Tanbin\Desktop\Bing_Hybrid_Framework1.xls"

DataTable.AddSheet "TestPlanInUFT"' to create a datatable during runtime datatable 

DataTable.AddSheet "TestStepsInUFT"

DataTable.ImportSheet FileLocation,"TestPlan","TestPlanInUFT"
DataTable.ImportSheet FileLocation,"TestCases","TestStepsInUFT"

TotalScenarios = DataTable.GetSheet("TestPlanInUFT").GetRowCount()
TotalSteps = DataTable.GetSheet("TestStepsInUFT").GetRowCount()

For i = 1  To TotalScenarios
		DataTable.GetSheet("TestPlanInUFT").SetCurrentRow(i)
		TSID = int(DataTable.value("TS_ID","TestPlanInUFT"))
        ExecuteKey = datatable.Value("Execution_Flag","TestPlanInUFT")

    If ExecuteKey = "Y" Then
    	
    For j = 1 To TotalSteps
    	DataTable.GetSheet("TestStepsInUFT").SetCurrentRow(j)
    	TCID = int(DataTable.value("TC_ID","TestStepsInUFT"))
    	If TSID = TCID Then
    		ActionKey = datatable.Value("Keyword","TestStepsInUFT")
    	    ObjId = datatable.Value("Obj_Info","TestStepsInUFT")
    	    ObjInput = datatable.Value("Obj_Input","TestStepsInUFT")
    	    Call ActionTaker(ActionKey, ObjId, ObjInput)
    	    Environment("TotalTC") = Environment("TotalTC") + 1
    	End If
    	
    Next
    
    
    If Environment("TotalTC") = Environment("TotalStep_Status")  Then
    	datatable.value("Test_Status","TestPlanInUFT") = "Passed"
    	
    	else
    	datatable.value("Test_Status","TestPlanInUFT") = "Failed"
    End If
       
    End If    
    
Next
DataTable.ExportSheet "C:\Users\Tanbin\Desktop\testResult2.xls","TestStepsInUFT"
