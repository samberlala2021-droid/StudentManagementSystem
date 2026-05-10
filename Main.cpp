// Topic: Student Management System.
// Group Member: [SamBat SamBer(LD)],[Oeng Sitong(SLD)], [Aing hangphinh], [Member 5], [Member 6]

// .gitignore — optional
// autopush.sh — just a git shortcut, ignore it
// Justfile — optional shortcut, don't need it
// readme.md — nice to have but not required

#include <iostream> //this is cin, cout
#include <vector> // lets you use vector<Student> - like a resizable array
#include <string> // lets you use string type properly
#include <algorithm> // gives you sort() and find_if()
#include <tabulate/table.hpp> // 3rd party library - draws those nice tables in terminal
#include <xlnt/xlnt.hpp> // 3rd party library - reads and writes Excel files
using namespace xlnt; // so write workbook instead of xlnt::workbook
using namespace tabulate;  // write Table instead of tabulate::Table
using namespace std;  // write cout instead of std::cout

// user class - store login info 
class User {
private:
    string username;
    string password;
    string role; // admin or user 
public:
    User(string username, string password, string role){
        this->username = username; 
        this->password = password; 
        this->role = role; 
    }
    string getUsername() { return username; }
    string getPassword() { return password; }
    string getRole() { return role; }
};

// student class 
class Student{
private:
    int id; 
    string name;
    int age;
    string gender;
    string major;  // added major 
    float gpa;     // added gpa 

public:
    Student(){ gpa = 0.0f; }
    Student(int id, string name, int age, string gender, string major, float gpa){   
        this->id = id; 
        this->name = name;
        this->age = age;
        this->gender = gender;
        this->major = major; 
        this->gpa = gpa; 
    }

    // getters 
    int getId() { return id; }
    string getName() { return name; }
    int getAge() { return age; }
    string getGender() { return gender; }
    string getMajor() { return major; }
    float getGpa() { return gpa; }

    // setters 
    void setId(int id) { this->id = id; }
    void setName(string name) { this->name = name; }
    void setAge(int age) { this->age = age; }
    void setGender(string gender) { this->gender = gender; }
    void setMajor(string major) { this->major = major; }
    void setGpa(float gpa) { this->gpa = gpa; }

    void input(){
        cout<<"Enter Student ID: ";cin>>id; 
        cout<<"Enter student name: ";cin.ignore();getline(cin,name); 
        cout<<"Enter student gender (M/F): ";cin>>gender; 
        cout<<"Enter student age: ";cin>>age; 
        cout<<"Enter student major: "; cin.ignore();getline(cin,major); 
        cout<<"Enter student GPA (0-4.0): ";cin>>gpa; 
    }

    // update student data 
    void update(){
        cout<<"Enter new name: ";cin.ignore();getline(cin,name); 
        cout<<"Enter new gender: ";cin>>gender; 
        cout<<"Enter new age: ";cin>>age; 
        cout<<"Enter new major: ";cin.ignore();getline(cin,major); 
        cout<<"Enter new GPA: ";cin>>gpa; 
    }
};

// hardcoded users for login 
vector<User> getUsers() {
    return {
        User("admin",   "admin123",  "admin"),
        User("teacher", "teach123",  "admin"),
        User("student", "student123","user"),
        User("guest",   "guest123",  "user")
    };
}

// helper functions for display
void clearScreen(){
#ifdef _WIN32
    system("cls");
#else
    system("clear");
#endif
}

void printLabel(string label){
    Table t;
    t.add_row({label});
    t[0].format().font_style({FontStyle::bold}).font_align({FontAlign::center});
    cout<<t<<endl;
}
void printMenuAsTable(string title, vector<string> items){
    Table t;
    t.add_row({title});
    t[0].format().font_style({FontStyle::bold}).font_align({FontAlign::center});
    for (auto& item : items) t.add_row({item});
    cout<<t<<endl;
}
void printStudentsAsTable(vector<Student> list){
    if(list.empty()){
        cout<<"No student records found.\n\n";
        return;
    }
    Table t;
    t.add_row({"ID", "Name", "Age", "Gender", "Major", "GPA"});
    t[0].format().font_style({FontStyle::bold}).font_align({FontAlign::center});
    for (auto& s : list) {
        t.add_row({
            to_string(s.getId()),
            s.getName(),
            to_string(s.getAge()),
            s.getGender(),
            s.getMajor(),
            to_string(s.getGpa()).substr(0, 4)
        });
    }
    cout<<t<<endl;
}

// save vector to excel file 
void saveStudentsToExcel(string filename, vector<Student>& list){
    workbook wb;
    auto ws = wb.active_sheet();
    ws.title("Students");

    ws.cell("A1").value("ID");
    ws.cell("B1").value("Name");
    ws.cell("C1").value("Age");
    ws.cell("D1").value("Gender");
    ws.cell("E1").value("Major");
    ws.cell("F1").value("GPA");

    int row = 2;
    for(auto& s : list){
        ws.cell("A" + to_string(row)).value(to_string(s.getId()));
        ws.cell("B" + to_string(row)).value(s.getName());
        ws.cell("C" + to_string(row)).value(to_string(s.getAge()));
        ws.cell("D" + to_string(row)).value(s.getGender());
        ws.cell("E" + to_string(row)).value(s.getMajor());
        ws.cell("F" + to_string(row)).value(to_string(s.getGpa()));
        row++;
    }
    wb.save(filename);
    cout<<"\n  Successfully saved data to file\n\n";
}

// read from excel and convert to student vector 
vector<Student> readDataFromExcel(string filename) {
    vector<Student> list;
    workbook wb;
    try {
        wb.load(filename);
    } catch (...) {
        return list;
    }
    worksheet ws = wb.active_sheet();
    for (auto row : ws.rows(false)) {
        if (row[0].to_string() == "ID") continue;
        try {
            int    id     = stoi(row[0].to_string());
            string name   = row[1].to_string();
            int    age    = stoi(row[2].to_string());
            string gender = row[3].to_string();
            string major  = row[4].to_string();
            float  gpa    = stof(row[5].to_string());
            list.push_back(Student(id, name, age, gender, major, gpa));
        } catch (...) {
            continue;
        }
    }
    return list;
}

// login function - returns logged in user 
User* login(vector<User>& users) {
    string username, password;
    int attempts = 3;
    while (attempts > 0) {
        clearScreen();
        printLabel(" Student Management System ");
        cout << "  Username : "; cin >> username;
        cout << "  Password : "; cin >> password;

        for (auto& u : users) {
            if (u.getUsername() == username && u.getPassword() == password) {
                cout << "\n  Login successful! Welcome, " << username
                     << " [" << u.getRole() << "]\n";
                return &u;
            }
        }
        attempts--;
        cout << "\n  Invalid credentials! Attempts left: " << attempts << "\n";
        cout << "  Press Enter to try again..."; cin.ignore(); cin.get();
    }
    cout << "\n  Too many failed attempts. Exiting.\n";
    return nullptr;
}

// show students with pagination - 5 per page 
void showWithPagination(vector<Student>& list) {
    int pageSize = 5;
    int total    = list.size();
    int page     = 0;
    int totalPages = (total + pageSize - 1) / pageSize;

    if (total == 0) { cout << "  No records found.\n\n"; return; }

    while (true) {
        clearScreen();
        printLabel(" All Students - Page " + to_string(page + 1) + " of " + to_string(totalPages));

        vector<Student> pageData;
        int start = page * pageSize;
        int end   = min(start + pageSize, total);
        for (int i = start; i < end; i++) pageData.push_back(list[i]);

        printStudentsAsTable(pageData);

        cout << "  [N] Next  [P] Previous  [Q] Quit\n  >> ";
        char c; cin >> c; c = toupper(c);
        if (c == 'N' && page < totalPages - 1) page++;
        else if (c == 'P' && page > 0) page--;
        else if (c == 'Q') break;
    }
}

// sort students by different fields 
void sortStudents(vector<Student>& list) {
    printLabel(" Sort Students ");
    vector<string> sortMenu = {
        "1. Sort by ID",
        "2. Sort by Name",
        "3. Sort by GPA (High to Low)",
        "4. Sort by Age"
    };
    printMenuAsTable(" Sort Options ", sortMenu);
    cout << "  >> Choose: "; int choice; cin >> choice;

    switch (choice) {
        case 1: sort(list.begin(), list.end(), [](Student& a, Student& b){ return a.getId() < b.getId(); }); break;
        case 2: sort(list.begin(), list.end(), [](Student& a, Student& b){ return a.getName() < b.getName(); }); break;
        case 3: sort(list.begin(), list.end(), [](Student& a, Student& b){ return a.getGpa() > b.getGpa(); }); break;
        case 4: sort(list.begin(), list.end(), [](Student& a, Student& b){ return a.getAge() < b.getAge(); }); break;
        default: cout << "  Invalid option.\n"; return;
    }
    cout << "\n  Sorted successfully!\n\n";
    printStudentsAsTable(list);
}

// search and filter students 
void searchStudents(vector<Student>& list) {
    printLabel(" Search / Filter Students ");
    vector<string> searchMenu = {
        "1. Search by Name",
        "2. Search by ID",
        "3. Filter by Gender",
        "4. Filter by Major"
    };
    printMenuAsTable(" Search Options ", searchMenu);
    cout << "  >> Choose: "; int choice; cin >> choice;

    vector<Student> results;

    if (choice == 1) {
        string keyword;
        cout << "  Enter name keyword: "; cin.ignore(); getline(cin, keyword);
        string kw = keyword;
        transform(kw.begin(), kw.end(), kw.begin(), ::tolower);
        for (auto& s : list) {
            string n = s.getName();
            transform(n.begin(), n.end(), n.begin(), ::tolower);
            if (n.find(kw) != string::npos) results.push_back(s);
        }
    } else if (choice == 2) {
        int id; cout << "  Enter ID: "; cin >> id;
        for (auto& s : list) if (s.getId() == id) results.push_back(s);
    } else if (choice == 3) {
        string gender; cout << "  Enter Gender (M/F): "; cin >> gender;
        for (auto& s : list) if (s.getGender() == gender) results.push_back(s);
    } else if (choice == 4) {
        string major; cout << "  Enter Major: "; cin.ignore(); getline(cin, major);
        string mj = major;
        transform(mj.begin(), mj.end(), mj.begin(), ::tolower);
        for (auto& s : list) {
            string m = s.getMajor();
            transform(m.begin(), m.end(), m.begin(), ::tolower);
            if (m.find(mj) != string::npos) results.push_back(s);
        }
    } else {
        cout << "  Invalid option.\n"; return;
    }

    cout << "\n  Found " << results.size() << " result(s):\n\n";
    printStudentsAsTable(results);
    cout << "  Press Enter to continue..."; cin.ignore(); cin.get();
}

// admin menu - full access 
void adminMenu(vector<Student>& list, string filename) {
    int option;
    do {
        clearScreen();
        printMenuAsTable(" Student Management System [ADMIN] ", {
            "1. Add New Student",
            "2. Edit Student",
            "3. Delete Student",
            "4. Search / Filter",
            "5. View All Students (Paginated)",
            "6. Sort Students",
            "7. Logout"
        });
        cout << "  >> Choose (1-7): "; cin >> option;

        switch (option) {
        case 1: {
            clearScreen();
            printLabel(" Add New Student ");
            Student s; s.input();
            list.push_back(s);
            saveStudentsToExcel(filename, list);
            cout << "  Student added successfully!\n";
            cout << "  Press Enter to continue..."; cin.ignore(); cin.get();
            break;
        }
        case 2: {
            clearScreen();
            printLabel(" Edit Student ");
            int id; cout << "  Enter Student ID to edit: "; cin >> id;
            auto it = find_if(list.begin(), list.end(), [&](Student& s){ return s.getId() == id; });
            if (it != list.end()) {
                cout << "  Student found! Enter new details:\n";
                it->update();
                saveStudentsToExcel(filename, list);
                cout << "  Student updated successfully!\n";
            } else {
                cout << "  Student with ID " << id << " not found!\n";
            }
            cout << "  Press Enter to continue..."; cin.ignore(); cin.get();
            break;
        }
        case 3: {
            clearScreen();
            printLabel(" Delete Student ");
            int id; cout << "  Enter Student ID to delete: "; cin >> id;
            auto it = find_if(list.begin(), list.end(), [&](Student& s){ return s.getId() == id; });
            if (it != list.end()) {
                list.erase(it);
                saveStudentsToExcel(filename, list);
                cout << "  Student deleted successfully!\n";
            } else {
                cout << "  Student with ID " << id << " not found!\n";
            }
            cout << "  Press Enter to continue..."; cin.ignore(); cin.get();
            break;
        }
        case 4:
            clearScreen();
            searchStudents(list);
            break;
        case 5:
            showWithPagination(list);
            break;
        case 6:
            clearScreen();
            sortStudents(list);
            cout << "  Press Enter to continue..."; cin.ignore(); cin.get();
            break;
        case 7:
            cout << "\n  Logging out...\n\n";
            break;
        default:
            cout << "  Invalid option!\n";
        }
    }while(option != 7);
}
// normal user menu - view and search only 
void userMenu(vector<Student>& list){
    int option;
    do{
        clearScreen();
        printMenuAsTable(" Student Management System [USER] ",{
            "1. Search / Filter Students",
            "2. View All Students (Paginated)",
            "3. Logout"
        });

        cout << "  >> Choose (1-3): "; cin >> option;

        switch(option){
        case 1:
            clearScreen();
            searchStudents(list);
            break;
        case 2:
            showWithPagination(list);
            break;
        case 3:
            cout << "\n  Logging out...\n\n";
            break;
        default:
            cout << "  Invalid option!\n";
        }
    }while(option != 3);
}
int main(){
    string filename = "studentdata.xlsx";
    vector<User> users = getUsers();

    while(true){
        User * loggedIn = login(users);
        if(!loggedIn) break;

        vector<Student> studentList = readDataFromExcel(filename);

        if(loggedIn->getRole() == "admin") {
            adminMenu(studentList, filename);
        }else{
            userMenu(studentList);
        }
        cout << "  Press Enter to go back to login or Ctrl+C to exit...\n";cin.ignore(); cin.get();
    }
    return 0;
}