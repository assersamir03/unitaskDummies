import xlsxwriter
data=[
{
"Name": "Fatima Hassan",
"Department": 5,
"Email": "fatima.hassan@btu.edu.eg",
"Semester": 2
},

{
"Name": "Omar Ali",
"Department": 6,
"Email": "omar.ali@btu.edu.eg",
"Semester": 1
},

{
"Name": "Lina Mansour",
"Department": 1,
"Email": "lina.mansour@btu.edu.eg",
"Semester": 4
},

{
"Name": "Ahmed Ibrahim",
"Department": 2,
"Email": "ahmed.ibrahim@btu.edu.eg",
"Semester": 3
},

{
"Name": "Yara Mahmoud",
"Department": 5,
"Email": "yara.mahmoud@btu.edu.eg",
"Semester": 2
},

{
"Name": "Nadia Salah",
"Department": 6,
"Email": "nadia.salah@btu.edu.eg",
"Semester": 4
},

{
"Name": "Khaled Samir",
"Department": 1,
"Email": "khaled.samir@btu.edu.eg",
"Semester": 3
},

{
"Name": "Sara Mansour",
"Department": 2,
"Email": "sara.mansour@btu.edu.eg",
"Semester": 1
},

{
"Name": "Mona Abdullah",
"Department": 5,
"Email": "mona.abdullah@btu.edu.eg",
"Semester": 4
},

{
"Name": "Hassan Mahmoud",
"Department": 6,
"Email": "hassan.mahmoud@btu.edu.eg",
"Semester": 2
},
{
"Name": "Sara Ibrahim",
"Department": 1,
"Email": "sara.ibrahim@btu.edu.eg",
"Semester": 3
},

{
"Name": "Ali Mahmoud",
"Department": 2,
"Email": "ali.mahmoud@btu.edu.eg",
"Semester": 4
},

{
"Name": "Lina Mansour",
"Department": 5,
"Email": "lina.mansour@btu.edu.eg",
"Semester": 2
},

{
"Name": "Omar Hassan",
"Department": 6,
"Email": "omar.hassan@btu.edu.eg",
"Semester": 1
},

{
"Name": "Yara Samir",
"Department": 1,
"Email": "yara.samir@btu.edu.eg",
"Semester": 4
},

{
"Name": "Ahmed Ibrahim",
"Department": 2,
"Email": "ahmed.ibrahim@btu.edu.eg",
"Semester": 3
},

{
"Name": "Nadia Salah",
"Department": 5,
"Email": "nadia.salah@btu.edu.eg",
"Semester": 2
},

{
"Name": "Khaled Mansour",
"Department": 6,
"Email": "khaled.mansour@btu.edu.eg",
"Semester": 1
},

{
"Name": "Mona Abdullah",
"Department": 1,
"Email": "mona.abdullah@btu.edu.eg",
"Semester": 3
},

{
"Name": "Hassan Mahmoud",
"Department": 2,
"Email": "hassan.mahmoud@btu.edu.eg",
"Semester": 4
},
{
"Name": "Ahmed Mansour",
"Department": 5,
"Email": "ahmed.mansour@btu.edu.eg",
"Semester": 1
},

{
"Name": "Laila Ibrahim",
"Department": 6,
"Email": "laila.ibrahim@btu.edu.eg",
"Semester": 3
},

{
"Name": "Ali Mahmoud",
"Department": 1,
"Email": "ali.mahmoud@btu.edu.eg",
"Semester": 2
},

{
"Name": "Yara Hassan",
"Department": 2,
"Email": "yara.hassan@btu.edu.eg",
"Semester": 4
},

{
"Name": "Omar Samir",
"Department": 5,
"Email": "omar.samir@btu.edu.eg",
"Semester": 1
},

{
"Name": "Sara Mansour",
"Department": 6,
"Email": "sara.mansour@btu.edu.eg",
"Semester": 3
},

{
"Name": "Khaled Ibrahim",
"Department": 1,
"Email": "khaled.ibrahim@btu.edu.eg",
"Semester": 2
},

{
"Name": "Nadia Ali",
"Department": 2,
"Email": "nadia.ali@btu.edu.eg",
"Semester": 4
},

{
"Name": "Mona Mahmoud",
"Department": 5,
"Email": "mona.mahmoud@btu.edu.eg",
"Semester": 1
},

{
"Name": "Hassan Mansour",
"Department": 6,
"Email": "hassan.mansour@btu.edu.eg",
"Semester": 3
},
{
"Name": "Sara Ibrahim",
"Department": 1,
"Email": "sara.ibrahim@btu.edu.eg",
"Semester": 4
},

{
"Name": "Ahmed Samir",
"Department": 2,
"Email": "ahmed.samir@btu.edu.eg",
"Semester": 1
},

{
"Name": "Lina Hassan",
"Department": 5,
"Email": "lina.hassan@btu.edu.eg",
"Semester": 3
},

{
"Name": "Omar Mansour",
"Department": 6,
"Email": "omar.mansour@btu.edu.eg",
"Semester": 2
},

{
"Name": "Yara Mahmoud",
"Department": 1,
"Email": "yara.mahmoud@btu.edu.eg",
"Semester": 4
},

{
"Name": "Nadia Ali",
"Department": 2,
"Email": "nadia.ali@btu.edu.eg",
"Semester": 1
},

{
"Name": "Khaled Samir",
"Department": 5,
"Email": "khaled.samir@btu.edu.eg",
"Semester": 3
},

{
"Name": "Mona Ibrahim",
"Department": 6,
"Email": "mona.ibrahim@btu.edu.eg",
"Semester": 2
},

{
"Name": "Hassan Mansour",
"Department": 1,
"Email": "hassan.mansour@btu.edu.eg",
"Semester": 4
},

{
"Name": "Laila Mahmoud",
"Department": 2,
"Email": "laila.mahmoud@btu.edu.eg",
"Semester": 1
}
]
import xlsxwriter
thefile = xlsxwriter.Workbook("theExelFile.xlsx")
thefileSheet = thefile.add_worksheet("firstSheet")
thefileSheet.write(0, 0, "#")
thefileSheet.write(0, 1, "#Name")
thefileSheet.write(0, 2, "Department")
thefileSheet.write(0, 3, "Email")
thefileSheet.write(0, 4, "Semester")

for i, entry in enumerate(data):
    thefileSheet.write(i + 1, 0, str(i))
    thefileSheet.write(i + 1, 1, entry["Name"])
    thefileSheet.write(i + 1, 2, entry["Department"])
    thefileSheet.write(i + 1, 3, entry["Email"])
    thefileSheet.write(i + 1, 4, entry["Semester"])

thefile.close()