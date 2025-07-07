import form
driver_path = r"D:\Google-maps-Contact-Details\Driver\chromedriver-win64\chromedriver-win64\chromedriver.exe"
Excel_Sheet_save_location = r"D:\Google-maps-Contact-Details\Excel\Save\ "


gm = form.GoogleMaps(Title="Robot Companies",Location="Saitama",Driver_Path=driver_path,Save_Excel=Excel_Sheet_save_location)
gm.search_text()
# Manpower Consultants Nepal