#include <iostream>
#include <vector>
#include <tabulate/table.hpp>
#include <xlnt/xlnt.hpp>
using namespace xlnt; 
using namespace tabulate;
using namespace std;
class Student
{
private:
    int id; 
    string name;
    int age;
    string gender;

public:
    Student(){}
    Student(int id , string name, int age, string gender)
    {   
        this->id  = id; 
        this->name = name;
        this->age = age;
        this->gender = gender;
    }

    string getGender() { return gender; }
    void setGender(string gender)
    {
        this->gender = gender;
    }
    string getName()
    {
        return name;
    }
    int getAge()
    {
        return age;
    }
    void setName(string name)
    {
        this->name = name;
    }
    void setAge(int age)
    {
        this->age = age;
    }

    int getId(){return id; }
    void setId(int id ){this->id = id; }
    // update 
    // void update(); 
    void update(){
        cout<<"Enter studnet name: "; 
        cin.ignore(); 
        getline(cin,name); 
        cout<<"Enter student gender: "; 
        cin>>gender; 
        cout<<"Enter student age: "; 
        cin>>age;
    }
    void input(){
        cout<<"Enter Student ID: "; 
        cin>>id; 
        cout<<"Enter studnet name: "; 
        cin.ignore(); 
        getline(cin,name); 
        cout<<"Enter student gender: "; 
        cin>>gender; 
        cout<<"Enter student age: "; 
        cin>>age; 
    }
};

void printStudentsAsTable(vector<Student> studentList)
{
    Table table;
    table.add_row({"ID","Name", "Age", "Gender"});
    for (auto student : studentList)
    {
        table.add_row({
            to_string(student.getId()),
            student.getName(),
            to_string(student.getAge()),
            student.getGender()}
        );
    }
    // on first row , make it bold
    table[0].format()
    .font_style({FontStyle::bold})
    .font_align({FontAlign::center});
    cout << table << endl;
}
void printMenuAsTable(vector<string> menuList)
{
    Table table;
    table.add_row({" Student Console App"});
    for (auto item : menuList)
    {
        table.add_row({item});
    }
    // on first row , make it bold
    table[0].format().font_style({FontStyle::bold}).font_align({FontAlign::center});
    cout << table << endl;
}

void printLabel(string label)
{
    Table table;
    table.add_row({label});
    table[0].format().font_style({FontStyle::bold})
    .font_align({FontAlign::center});
    cout << table << endl;
}


// Save vector to excel file 
void saveStudentsToExcel(
    string filename, 
    vector<Student> studentList 
){

    // select active sheet , give the sheet name 
    workbook wb; 
    auto ws = wb.active_sheet(); 
    ws.title("Sheet1"); 

    // add data 
    // table header 
    ws.cell("A1").value("ID");
    ws.cell("B1").value("Name"); 
    ws.cell("C1" ).value("Age");
    ws.cell("D1" ).value("Gender");

    int row= 2; 
    for(auto student: studentList){
        ws.cell("A"+to_string(row))
            .value(to_string(student.getId()));
        ws.cell("B"+to_string(row)).value(student.getName()); 
        ws.cell("C"+to_string(row) ).value(to_string(student.getAge()));
        ws.cell("D"+to_string(row) ).value(student.getGender());
        row++; 
    }

    // save file 
    wb.save(filename); 
    cout<<" ✅ Successfully saved data to file"<<endl; 



}
// read from excel and convert it to student vector 
vector<Student> readDataFromExcel(string& filename){
    vector<Student> students; 
    workbook wb; 
     try{
         wb.load(filename); 
     }catch(...){
         cout<<"Excel file cannot be open for reading!! ❌"<<endl; 
         return students; 
     }
   
    // open sheet , loop through row and convert to vector 
   worksheet ws = wb.active_sheet(); // Sheet1 
    for( auto row : ws.rows(false)){
        // skip table header 
        if( row[0].to_string()== "ID") continue; 

        int id ; 
        string name; 
        int age; 
        string gender; 

        // stoi = string to integer 
        id= stoi(row[0].to_string()); 
        name = row[1].to_string(); 
        age = stoi(row[2].to_string()); 
        gender = row[3].to_string();  

        Student student(id,name,age,gender); 
        students.push_back(student); 
    }
    return students; 
}

int main()
{
    system("clear");
    vector<Student> studentList;
    string filename = "studentdata.xlsx";
    vector<string> menuList = {
        "1. Add new Student",
        "2. Edit Student Data ",
        "3. Search Student Data ",
        "4. Delete Student Data ",
        "5. Show All Student Data ",
        "6. Exit"};
    int option;
    do
    {
        studentList = readDataFromExcel(filename); 
        printMenuAsTable(menuList);
        cout << ">> CHoose optioon(1-6): ";
        cin >> option;
        switch (option)
        {
        case 1:
        {
            printLabel("Add New Student"); 
            Student newStudent; 
            newStudent.input(); 
            studentList.push_back(newStudent); 
            cout<<" ✅ Successfully add new student! "<<endl; 
            saveStudentsToExcel(filename,studentList); 

        }
            break;
        case 2:
        {
            printLabel("Edit Student Data "); 
            int studentId; 
            // find_if 
            cout<<"Enter student id to update: "; 
            cin>>studentId; 

            auto result = find_if(
                studentList.begin(), 
                studentList.end(),
                [&](Student& student){
                    return student.getId()==studentId;
                }
            );

            if(result!=studentList.end()){
                // Result found 
                cout<<"Student with = "
                    <<studentId<<" found ✅"<<endl; 
                // update the data 
                Student student  = *result; 
                student.update(); 
                // update the values of the list 
                *result = student; 
            }else {
                // result not found 
                cout<<"Student with the ID ="<<studentId
                    <<" doesn't exit ❌"<<endl; 

            }
        }
            break;
        case 3:
            break;
        case 4:{
            printLabel("Delete Student Data "); 
            cout<<"Enter ID to delete : "; 
            int studentId; 
            cin>>studentId; 

             auto result = find_if(
                studentList.begin(), 
                studentList.end(),
                [&](Student& student){
                    return student.getId()==studentId;
                }
            );

            if(result!=studentList.end()){
                // delete here 
                studentList.erase(result); 
                cout<<" ✅ Successfully Delete Data ! "<<endl; 
            }else {
                cout<<"Student with ID = "<<studentId
                <<" doesn't exist ! ❌"
                    <<endl; 
            }
        }
            break;
        case 5:
            sort(
                studentList.begin(), 
                studentList.end() , 
                [](Student& s1, Student& s2){
                    return s1.getId() < s2.getId();
                }
            ); 
            printStudentsAsTable(studentList);
            break;
        case 6:
            cout << "Exit from the program" << endl;
            break;
        }
    } while (option != 6);
    return 0;
}