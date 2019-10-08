# Tables Data Extraction (Creo Automation)
The program demonstrates the extraction of table values from a Creo drawing file using VB API

Preliminary Steps:
1. Ensure you have installed VB API while installing creo (ie choose customize and select the Visual Basic API)
2. Add the PRO_COMM_MSG_EXE enviroment variable and point it to the executable file pro_comm_msg.exe. (in my case, C:\Program Files\PTC\Creo 5.0.0.0\Common Files\x86e_win64\obj\pro_comm_msg.exe)
3. Run "vb_api_register.bat" (in my case found at, C:\Program Files\PTC\Creo 5.0.0.0\Parametric\bin\vb_api_register.bat)


Add Reference:
1. After creating a new project, Click Project --> Add Reference --> COM and select 'Creo VB API Type Library for Creo Parametric 5.0.0.0' and click ok 


Software Versions:
1. Creo Parametric 5.0.0.0
2. Visual Studio Community 2017
3. Dot Net Framework 4.5
