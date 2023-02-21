# Excel2JSON
VBA macros for exporting an excel sheet to Json format

Simple macros that will export a excel sheet like this:

![image](https://user-images.githubusercontent.com/1207326/220346762-061d8986-acd8-43f4-82cc-7bd61f120753.png)

to a json file called sheet_name.json with this structure:

![image](https://user-images.githubusercontent.com/1207326/220351578-fe247cc9-0c98-40a7-96a0-166a2a1c1b21.png)

By default, the file will be created in C:\Users\Usuario\Desktop. To change this, it has to be done inside the macros code:

![image](https://user-images.githubusercontent.com/1207326/220351886-6e1cc7f8-ccd6-4dac-996f-6f605ded817f.png)


- In orther to work, it is required to:
  - Have a first row with the column names
  - Have a second row with the data types: 
    - 'str' for strings
    - 'num' for ints, floats and booleans
    - 'arr_str' for arrays of strings
    - 'arr_num' for arrays of ints, floats and booleans
  - Have Array data values separated by coma and NO spaces

By default, the macros is run using ctrl + E
