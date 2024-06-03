package example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Data_table {
    private List<Student> students;
    private List<Teacher> teachers;
    private List<Subject> subjects;

    public Data_table() {
        this.students = new ArrayList<>();
        this.teachers = new ArrayList<>();
        this.subjects = new ArrayList<>();
    }

    public void add_student(Student student) {
        students.add(student);
    }

    public void add_teacher(Teacher teacher) {
        teachers.add(teacher);
    }

    public void add_subject(Subject subject) {
        subjects.add(subject);
    }


    public void readFromExcel(String file_path) {
        try {
            FileInputStream excelFile = new FileInputStream(new File(file_path));
            Workbook workbook = new XSSFWorkbook(excelFile);

            // первый лист
            Sheet studentsSheet = workbook.getSheetAt(0);
            for (int i = 0; i < studentsSheet.getLastRowNum(); i++) {
                Row studentNameRow = studentsSheet.getRow(i);
                Row subjectRow = studentsSheet.getRow(i );
                Row markRow = studentsSheet.getRow(i);

                String studentName = studentNameRow.getCell(0).getStringCellValue();
                Map<Subject, Integer> marksMap = new HashMap<>();
                String subjectName = subjectRow.getCell(1).getStringCellValue();
                int mark = (int) markRow.getCell(2).getNumericCellValue();
                Subject subject = new Subject(subjectName);
                marksMap.put(subject, mark);

                Student student = new Student(studentName, marksMap);
                add_student(student);


            }

            // второй лист
            Sheet TeacherSheet = workbook.getSheetAt(1);
            for (int i = 0; i < TeacherSheet.getLastRowNum(); i++) {
                Row teacherNameRow = TeacherSheet.getRow(i);
                Row subjectRow = TeacherSheet.getRow(i);

                String teacherName = teacherNameRow.getCell(0).getStringCellValue();
                String subjectName = subjectRow.getCell(1).getStringCellValue();
                Subject subject = new Subject(subjectName);
                Teacher teacher = new Teacher(teacherName, subject);
                add_teacher(teacher);

            }

            workbook.close();
            excelFile.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public JSONObject to_Json() {
        JSONObject jsonObject = new JSONObject();
        JSONArray studentsArray = new JSONArray();
        JSONArray teachersArray = new JSONArray();
        JSONArray subjectsArray = new JSONArray();

        for (Student student : students) {
            JSONObject studentObject = new JSONObject();
            studentObject.put("имя", student.getName());

            JSONArray marksArray = new JSONArray();
            for (Map.Entry<Subject, Integer> entry : student.getMarks().entrySet()) {
                JSONObject markObject = new JSONObject();
                markObject.put("предмет", entry.getKey().getName());
                markObject.put("средняя оценка", entry.getValue());
                marksArray.put(markObject);
            }
            studentObject.put("средние оценки", marksArray);

            studentsArray.put(studentObject);
        }

        for (Teacher teacher : teachers) {
            JSONObject teacherObject = new JSONObject();
            teacherObject.put("имя", teacher.getName());
            teacherObject.put("предмет", teacher.getSubject().getName());
            teachersArray.put(teacherObject);
        }

        for (Subject subject : subjects) {
            JSONObject subjectObject = new JSONObject();
            subjectObject.put("имя", subject.getName());
            subjectsArray.put(subjectObject);
        }

        jsonObject.put("студенты", studentsArray);
        jsonObject.put("преподы", teachersArray);
        jsonObject.put("предметы", subjectsArray);

        return jsonObject;
    }

    public static void writeJsonToFile(JSONObject jsonObject, String filePath) {
        try (FileWriter file = new FileWriter(filePath)) {
            file.write(jsonObject.toString());
            System.out.println("JSON успешно записан в файл: " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static JSONObject readJsonFromFile(String filePath) {
        try (FileReader reader = new FileReader(filePath)) {
            StringBuilder jsonString = new StringBuilder();
            int c;
            while ((c = reader.read()) != -1) {
                jsonString.append((char) c);
            }
            return new JSONObject(jsonString.toString());
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }


    public void writeToExcelFromJson(JSONObject json, String filePath) {
        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet studentsSheet = workbook.createSheet("Студенты");
            Sheet teachersSheet = workbook.createSheet("Преподы");

            JSONArray studentsArray = json.getJSONArray("студенты");
            int rowCount = 0;
            for (int i = 0; i < studentsArray.length(); i++) {
                JSONObject studentObject = studentsArray.getJSONObject(i);
                String name = studentObject.getString("имя");

                JSONArray marksArray = studentObject.getJSONArray("средние оценки");
                Row row = studentsSheet.createRow(rowCount++);
                int columnCount = 0;
                row.createCell(columnCount++).setCellValue(name);

                for (int j = 0; j < marksArray.length(); j++) {
                    JSONObject markObject = marksArray.getJSONObject(j);
                    String subjectName = markObject.getString("предмет");
                    int mark = markObject.getInt("средняя оценка");

                    row.createCell(columnCount++).setCellValue(subjectName);
                    row.createCell(columnCount++).setCellValue(mark);
                }
            }

            JSONArray teachersArray = json.getJSONArray("преподы");
            rowCount = 0;
            for (int i = 0; i < teachersArray.length(); i++) {
                JSONObject teacherObject = teachersArray.getJSONObject(i);
                String name = teacherObject.getString("имя");
                String subject = teacherObject.getString("предмет");

                Row row = teachersSheet.createRow(rowCount++);
                int columnCount = 0;
                row.createCell(columnCount++).setCellValue(name);
                row.createCell(columnCount++).setCellValue(subject);
            }

            FileOutputStream outputStream = new FileOutputStream(filePath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
            System.out.println("Данные успешно записаны в файл Excel: " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}