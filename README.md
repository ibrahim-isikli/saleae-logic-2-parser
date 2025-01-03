# saleae-logic-2-parser
This script parses Saleae Logic 2 output files (.csv) to process data packets byte by byte. Users can configure a threshold time, which is the time difference between two consecutive start times (start_time1 and start_time2), to efficiently identify and analyze data packets in the captured signal.

# Prerequisites

- Saleae Logic 2 output file .csv :

 ![single_analyzer_export](https://github.com/user-attachments/assets/18965817-b1fc-40b8-8e0f-0386883cf837)
 ![image](https://github.com/user-attachments/assets/bb06500b-28b0-4277-867a-cfaa34968bff)

- The .csv file will look like this:
  ![image](https://github.com/user-attachments/assets/868b117e-bd9b-47e0-a26b-f5d70e095efa)

  
- Save the file with .csv extension as .xlsm extension.
- Go to developer tab in excel. Then click on the Visual Basic section and create a module.
  ![image](https://github.com/user-attachments/assets/973a3e29-8001-4023-adc7-adfbaf51ecc0)

-Copy the script you want here

# Running the script

- logic2_parse_data_packets.vbs
- ![image](https://github.com/user-attachments/assets/10631cb5-60ab-4069-ba3b-ab6c0a5628c4)

- If you want to sort one under the other logic2_parse_data_packets.vbs then logic2_group_packages_in_rows.vbs
You can use the module.
![image](https://github.com/user-attachments/assets/ea810a8e-62ea-4774-8c3a-03c55c930686)




  
  
