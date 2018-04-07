# fluffy-chainsaw :see_no_evil: 
Online VBScript Repo :)

## Instructiones to proceed further :)
* If you know how to do Git clone it locally, or just download the folder :)
* Navigate to the `simple_to_check_environment` folder and run the files using IE, if it runs then proceed further.
* Navigate to the `WIP` folder and open the EDITHERE.html file using text editor, and change 
	* line 18 QNN to your Question number, 
	* line 23 buttonNN to your Question number.
* If you have any files to manipulate put it in the files folder
* Edit your script QNN.vbs in VBScript folder (Try to put everything into that one function pls)
* Have fun >.<'

## Questions Finished So far :)

- [X] Q01 Team_15 - Swathi Manoharan, Nivedha Ramesh
- [ ] ~~Q02 Team_03 - Thamilarasan Palani, Vishal Chacko~~
- [ ] ~~Q03 Team_06 - Bala Subramaniam Arunachalam, Balajivenkateswaran Sundarajan~~
- [X] Q04 Team_12 - Jayadevi Durairajulu, Jagadish Kumar Muthalur Chandrasekar
- [X] Q05 Team_14 - Munish Kumar Karthick Muni, Kumara Raja Subramani
- [ ] ~~Q06 Team_01 - Nivedha Ganeshan, Christus Colman Vincent Paul Arokiadass~~
- [X] Q07 Team_13 - Ashlin Jenefer Christopher, Meenatchi Ramesh
- [X] Q08 Team_09 - Sharika Sasikumar, Aravind Sagar Sellapandian
- [ ] ~~Q09 Team_04 - Kaarthikaa Manimaran, Srinivasan Subramaniam~~
- [X] Q10 Team_11 - Sadhana Sundar, Sairaj Raju
- [X] Q11 Team_05 - Mary Steffin Avila Rayen, Joseph Rajesh Joseph Antony Raj
- [X] Q12 Team_08 - Venkatesh Parameswaran, Karthik Subramanian
- [X] Q13 Team_02 - Maheswari Mohana Sundaram, Mohan Rajasekaran
- [X] Q14 Team_16 - Sudhankarthick Jaisankar, Sasikumar Kalimuthu
- [X] Q15 Team_07 - Mohamed Aejaz Mohamed Arifullam, Sameera Sareen Mohamed Ibrahim
- [X] Q16 Team_10 - Vignesh Kumar Sundararaj, Boopalan Natarajan

## :heavy_check_mark: Task list  

- [x] Create repo and edit Readme
- [x] Create simple HTML to check environment
- [x] Create Basic UI that IE10 supports
- [x] Enter the questions
- [x] Started adding individual scripts
- [X] 50% completion :)
- [ ] Integrate the answers
- [ ] Tweak it a little

## :pushpin: Note: 
> Download the repo and run the index.html in IE browser.
During run time, allow the blocked content (Scripts/ActiveX) in order to execute the scripts. Even after enabling the 'Active X', if it doesn't work, then try the following :)

* Open the IE browser (I have IE11)
* Go to 'Developer tool' options (CTRL + U)
* Click on 'DOM Explorer' tab
* On the right side, change IE browser version to 10 or lower :)

### :paperclip: Further notes 
> Some stuff to keep in mind :)
* Full Page JS compatibility IE9 and above
* VBscript compatibility IE10 and below
* Dim/Const did not work in IE9 (?)
* Added external vbs file... wondering what is the point of the 
* `Set oShell = CreateObject ("WScript.Shell")` :person_frowning:
* `Set WshShell = CreateObject("WScript.Shell")` :person_with_pouting_face:
* Did splits in scripting, trying to layout basic skeletons :confused:

## Resources Reffered to :)
* https://msdn.microsoft.com/en-us/library/t0aew7h6.aspx
* https://www.stardock.com/products/desktopx/help/dev_guide_3a_scripting.htm
* https://msdn.microsoft.com/en-us/library/bb149081.aspx
* https://msdn.microsoft.com/en-us/library/office/ee814737(v=office.14).aspx
* https://www.voiceguide.com/vghelp/source/html/modvbs.htm
* https://www.devhut.net/category/ms-access/ms-access-excel-automation/
* and lots other that got lost due to frequent clearing history and cache for performance testing... *sigh*