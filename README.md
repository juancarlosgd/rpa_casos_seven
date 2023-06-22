# rpa_casos_seven
Repository for Workflow Automation RPA

## First steps
This RPA is designed to be used with Python in SEVEN ERP workflows. Combines the use of Selenium and Pywin32 for Windows OS operations. SEVEN is an ERP that uses an ISAPI web model running on top of internet explorer and under the hood it runs as a WIN32 library.

Main Required libraries:
- Selenium
- Pywin32
  
**other required libraries are in the file requirements.txt**

## Configuration
1. Download and configure IE driver path
2. Configure internet explorer (path)
3. create windows registry key **Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BFCACHE**

  For 64-bit Windows installations,
  HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\Main\FeatureControl
  Create key, FEATURE_BFCACHE, if not already present.
  Inside this key, create a DWORD value named iexplore.exe with the value of 0. Even if QWORD is suggested for 64-bit machines, create a DWORD.

4. parameterize configuration.ini file
