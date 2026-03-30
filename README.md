# TODOs
1. [x] Create a sheet called `Danh sách gửi mail`
2. [x] Clone data from original sheet to another
3. [x] Get sheet by name, not by using `getActiveSheet`# truc-lam-meditation-course-for-adult
4. [x] Filter duplicated names and date of birth for kids
5. [x] Dynamic content such as current year, target students..., etc coming from "Lưu trữ" 
6. [x] Apply dynamic content to the form so that the user only needs to fill data to the sheet "Lưu trữ"
7. [x] Have a function to get data from "Sao kê" sheet and update the "Danh sách gửi mail" sheet so that the "Đã chuyển khoản" column is updated properly. This function will get run every 4 hours
    - In "Sao kê" sheet, grab the content of transfer note which is a merged cell from multiple cells spanning from column Y to AF. It has the structure "<sender name> - <phone number> - <code>" for example "Đặng Hưng - 0375072848 - ts". Only take the data that has this structure
    - In "Sao kê" sheet, grab the transfer amount which is probably a string that is separate by comma in a merged cell of columns from AT to BA. 
    - After grabbing these data, process the string to extract the phone number. Use this phone number to loop over every row in "Danh sách gửi mail" sheet to find the matching row. If the phone number is found, update the "Đã chuyển khoản" column with letter "x". Highlight the row in "Sao kê" with green background color
8. [ ] Create passengers list for each bus. Even when there are many buses, they all must live in only one sheet. Detail implementation is in @item-8-passenger-list-implementation-plan.md file
9. [ ] Create a "In danh sách xe" menu option in "Khoá tu" menu to generate real google doc files based on group of columns in a "Danh sách xe" sheet. 
    - The group of columns in "Danh sách xe" sheet takes every 5 columns starting from column A. Then it exclude one column and take another group of 5 columns. It continues this pattern until it reaches the end of the sheet horizontally. For example A to E, skip F then start again with G to K, skip L and so on. 
    - Save it to "Danh sách xe" folder which id is provided for you
10. [ ] Update Code.js to adapt to many different types of sheets. Below is the list of sheets:
    10.1. 