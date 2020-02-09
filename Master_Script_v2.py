#!/usr/bin/env python
# coding: utf-8

# In[4]:


if __name__ == '__main__':
    import os, os.path
    import win32com.client
    import time
    import sys
    import pandas as pd
    import numpy

    #file locations
    fileStringStart = str("D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\Tools\\")
    fileStringEnd = str(".xlsm")
    controlWorkbook = str("D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\Control Sheet - Industry Review.xlsm")
    nameplateRankWorkbook = str("\\_CommonSlides\\Nameplate Rank Charts (M-M & YTD-YTD).xlsm")
    scorecardWorkbook = str("\\_CommonSlides\\Health Metrics Scorecard Tool.xlsm")
    modelPerformanceWorkbook = str("\\Nameplate+Model Performance Slide Tool 16x9 - ")
    nameplateShareWalkWorkbook = str("\\Nameplate Share Walk Chart Tool - Region + Segment - ")
    regionalOverviewWorkbook = str("\\Nameplate+Model Regional Overview Slide Tool - ")

    #function to open workbooks
    def openWorkbook(fileLocation):
        if os.path.exists(fileLocation):
            global xl, wb
            xl=win32com.client.Dispatch("Excel.Application")
            wb = xl.Workbooks.Open(os.path.abspath(fileLocation))
            xl.Application.Visible = True   

    #function to close workbook and delete excel.exe in task manager
    def closeAndSaveWorkbook():
        wb.Close(SaveChanges=1)
        del  globals()['xl']
        time.sleep(1)

    # Common Slides
    def common_slides(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        openWorkbook(controlWorkbook)
        xl.Run("OpenPPT")
        closeAndSaveWorkbook()
        #M/M & YTD/YTD Rank Slide
        openWorkbook(fileStringStart + nameplateRankWorkbook)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("PPTAll")
        closeAndSaveWorkbook()
        #Scorecard Slides
        openWorkbook(fileStringStart + scorecardWorkbook)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Save CommonSlides
        openWorkbook(controlWorkbook)
        xl.Run("SaveCommonSlides")
        closeAndSaveWorkbook()

    # Kia Industry Review
    def Kia_IR(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Kia"
        #Open template
        openWorkbook(controlWorkbook)
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("NationalSlides")
        xl.Run("RegionalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Overview Slides
        openWorkbook(fileStringStart + nameplate + regionalOverviewWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPTNameplate")
        xl.Run("AutoPPTModels")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook) 
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("Sort" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT"+ nameplate)
        closeAndSaveWorkbook()

    # Hyundai Industry Review
    def Hyundai_IR(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Hyundai"
        #Open template
        openWorkbook(controlWorkbook)
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("NationalSlides")
        xl.Run("RegionalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Overview Slides
        openWorkbook(fileStringStart + nameplate + regionalOverviewWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPTNameplate")
        xl.Run("AutoPPTModels")
        closeAndSaveWorkbook()
        #Add Genesis Slides
        openWorkbook(fileStringStart + "Genesis" + modelPerformanceWorkbook + "Genesis" + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("NationalSlides")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook) 
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("Sort" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT"+ nameplate)
        closeAndSaveWorkbook()
     
    # Honda Industry Review
    def Honda_IR(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Honda"
        #Open template
        openWorkbook(controlWorkbook)
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)  
        xl.Run("NationalSlides")
        xl.Run("RegionalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("natCore")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook) 
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("Sort" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT"+ nameplate)
        closeAndSaveWorkbook()

    # Acura Industry Review
    def Acura_IR(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Acura"
        #Open template
        openWorkbook(controlWorkbook)
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)   
        xl.Run("NationalSlides")
        xl.Run("RegionalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("natCore")
        closeAndSaveWorkbook()
        #NSX Slide
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + " NSX" + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)   
        xl.Run("AutoPPTModel")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook) 
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("Sort" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT"+ nameplate)
        closeAndSaveWorkbook()

    # Mitsubishi Industry Review
    def Mitsu_IR(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Mitsubishi"
        #Open template
        openWorkbook(controlWorkbook)
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)   
        xl.Run("NationalSlides")
        xl.Run("RegionalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Overview Slides
        openWorkbook(fileStringStart + nameplate + regionalOverviewWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPTNameplate")
        xl.Run("AutoPPTModels")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook) 
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("Sort" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT"+ nameplate)
        closeAndSaveWorkbook()

    # Mazda Industry Review
    def Mazda_IR(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Mazda"
        #Open template
        openWorkbook(controlWorkbook)
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)   
        xl.Run("NationalSlides")
        xl.Run("RegionalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Overview Slides
        openWorkbook(fileStringStart + nameplate + regionalOverviewWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPTNameplate")
        xl.Run("AutoPPTModels")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook) 
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("Sort" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT"+ nameplate)
        closeAndSaveWorkbook()
    
    # Nissan Industry Review
    def Nissan_IR(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Nissan"
        #Open template
        openWorkbook(controlWorkbook)
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)  
        xl.Run("SinglePPTNameplate")
        xl.Run("PowerpointDeckAll")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("regional")
        xl.Run("segment")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Overview Slides
        openWorkbook(fileStringStart + nameplate + regionalOverviewWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Overview Slides (No Fullsize Pickups)
        openWorkbook(fileStringStart + nameplate + regionalOverviewWorkbook + nameplate + "_no_FSPU" + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Slides
        openWorkbook(fileStringStart + nameplate + "\\Nameplate+Model Regional Slide Tool 16x9 - " + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("AutoRegCopyPPT")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook) 
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("Sort" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT"+ nameplate)
        closeAndSaveWorkbook()
        
    #Hyundai IP
    def Hyundai_IP(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Hyundai"
        controlWorkbook = str("D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\Control Sheet - Incentive Planning.xlsm")
        forecastWorkbook = "D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\Forecast\\" + nameplate + " Forecast.xlsm"
        fileStringDSS = "D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\DSS\\" + nameplate + "\\"
        #Open PPT
        openWorkbook(controlWorkbook)
        xl.Run("OpenPPT")
        closeAndSaveWorkbook()
        #M/M & YTD/YTD Rank Slide
        openWorkbook(fileStringStart + nameplateRankWorkbook)    
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("PPTAll")
        closeAndSaveWorkbook()
        #Save CommonSlides
        openWorkbook(controlWorkbook)    
        xl.Run("SaveCommonSlides")
        closeAndSaveWorkbook()
        #Open Template
        openWorkbook(controlWorkbook)    
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + " IP" + fileStringEnd)    
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)    
        xl.Run("NationalSlides")
        xl.Run("RegionalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + " IP" + fileStringEnd)  
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)    
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Overview Slides
        openWorkbook(fileStringStart + nameplate + regionalOverviewWorkbook + nameplate + " IP" + fileStringEnd)   
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)    
        xl.Run("AutoPPTNameplate")
        xl.Run("AutoPPTModels")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook)    
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("SortStandardSlides" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT" + nameplate)
        closeAndSaveWorkbook()
    
    #Kia IP
    def Kia_IP(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Kia"
        controlWorkbook = str("D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\Control Sheet - Incentive Planning.xlsm")
        forecastWorkbook = "D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\Forecast\\" + nameplate + " Forecast.xlsm"
        fileStringDSS = "D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\DSS\\" + nameplate + "\\"
        #Open PPT
        openWorkbook(controlWorkbook)
        xl.Run("OpenPPT")
        closeAndSaveWorkbook()
        #M/M & YTD/YTD Rank Slide
        openWorkbook(fileStringStart + nameplateRankWorkbook)    
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("PPTAll")
        closeAndSaveWorkbook()
        #Save CommonSlides
        openWorkbook(controlWorkbook)    
        xl.Run("SaveCommonSlides")
        closeAndSaveWorkbook()
        #Open Template
        openWorkbook(controlWorkbook)    
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + " IP" + fileStringEnd)    
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)    
        xl.Run("NationalSlides")
        xl.Run("RegionalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + " IP" + fileStringEnd)  
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)    
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Nameplate+Model Regional Overview Slides
        openWorkbook(fileStringStart + nameplate + regionalOverviewWorkbook + nameplate + " IP" + fileStringEnd)   
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)    
        xl.Run("AutoPPTNameplate")
        xl.Run("AutoPPTModels")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook)    
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("SortStandardSlides" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT" + nameplate)
        closeAndSaveWorkbook()
        
        ### Add Forecast Slides
        #Open Output PPT
        openWorkbook(controlWorkbook)    
        xl.Run("OpenOutput" + nameplate)
        closeAndSaveWorkbook()
        #Add Forecast Slides
        openWorkbook(forecastWorkbook)    
        xl.Run("AutoPPT")
        closeAndSaveWorkbook()
        #Sort Forecast Slides
        openWorkbook(controlWorkbook)    
        xl.Run("SortForecastSlides" + nameplate)
        xl.Run("SavePPTWithForecast" + nameplate)
        closeAndSaveWorkbook()
        
        ### Add DSS Slides
        #Open Output with Forecast Slides PPT
        openWorkbook(controlWorkbook)    
        xl.Run("OpenOutputWithForecast" + nameplate)
        closeAndSaveWorkbook()
        #Grab dss file names in Control Workbook
        df = pd.read_excel(controlWorkbook, sheet_name=nameplate)
        df_slice = df.iloc[25:46,2:5]
        total_model_count = df_slice.iloc[0,0]
        fileListDSS = []
        for x in df_slice.iloc[1:,1]:
            if isinstance(x, str):
                fileListDSS.append(x)
        #Add DSS Slides
        for i in fileListDSS:
            filePath = fileStringDSS + i + '.xlsm'
            openWorkbook(filePath)
            xl.Run("AutoPPTModels")
            xl.Run("AutoPPTSummary")
            closeAndSaveWorkbook()
        #Sort DSS Slides
        openWorkbook(controlWorkbook)    
        xl.Run("SortDSSSlides" + nameplate)
        xl.Run("ReOrderDSSSlides" + nameplate)
        xl.Run("SavePPTWithForecastAndDSS" + nameplate)
        closeAndSaveWorkbook()

    def Mitsu_IP(report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast):
        nameplate = "Mitsubishi"
        controlWorkbook = str("D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\Control Sheet - Incentive Planning.xlsm")
        forecastWorkbook = "D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\Forecast\\" + nameplate + " Forecast.xlsm"
        fileStringDSS = "D:\\Users\\bryant.vu\\Desktop\\Monthly Automation\\DSS\\" + nameplate + "\\"
        #Open PPT
        openWorkbook(controlWorkbook)    
        xl.Run("OpenPPT")
        closeAndSaveWorkbook()
        #M/M & YTD/YTD Rank Slide
        openWorkbook(fileStringStart + nameplateRankWorkbook)    
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("PPTAll")
        closeAndSaveWorkbook()
        #Save CommonSlides
        openWorkbook(controlWorkbook)    
        xl.Run("SaveCommonSlides")
        closeAndSaveWorkbook()
        #Open Template
        openWorkbook(controlWorkbook)    
        xl.Run("OpenTemplate" + nameplate)
        closeAndSaveWorkbook()
        #Nameplate+Model Peformance Slides
        openWorkbook(fileStringStart + nameplate + modelPerformanceWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("NationalSlides")
        closeAndSaveWorkbook()
        #Nameplate Walk Charts Slides
        openWorkbook(fileStringStart + nameplate + nameplateShareWalkWorkbook + nameplate + fileStringEnd)
        xl.Run("update_slicer", report_date, refresh_cube, monthend_or_MTD, share_date, spend_date, forecast)
        xl.Run("natCore")
        xl.Run("regional")
        closeAndSaveWorkbook()
        #Add Common Slides, Section Titles & Sort Slides
        openWorkbook(controlWorkbook)    
        xl.Run("CopyFromCommonSlides" + nameplate)
        xl.Run("AddPPTSections" + nameplate)
        xl.Run("SortStandardSlides" + nameplate)
        xl.Run("UpdateDeckCover" + nameplate)
        xl.Run("SavePPT" + nameplate)
        closeAndSaveWorkbook()

    #Run functions
    if sys.argv[1] == 'Industry Review (all)':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Acura_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Honda_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Hyundai_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Kia_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Mazda_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Mitsu_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Nissan_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Acura_IR':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Acura_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Honda_IR':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Honda_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Hyundai_IR':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Hyundai_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Kia_IR':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Kia_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Mazda_IR':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Mazda_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Mitsu_IR':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Mitsu_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Nissan_IR':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Nissan_IR(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Hyundai_IP':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Hyundai_IP(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Kia_IP':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Kia_IP(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'Mitsu_IP':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Mitsu_IP(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
    elif sys.argv[1] == 'IP Decks (all)':
        common_slides(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Kia_IP(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Hyundai_IP(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        Mitsu_IP(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])

