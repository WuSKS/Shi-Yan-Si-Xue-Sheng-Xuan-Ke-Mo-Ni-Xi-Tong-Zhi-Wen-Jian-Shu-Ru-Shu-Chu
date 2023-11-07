package CeShi;


import java.io.FileOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import XuanKe.Course;
import XuanKe.Students;
import XuanKe.Teachers;


public class Test {
	
	public static void main(String[] args) {
        // 实例化课程对象
        Course course1 = new Course(1, "\t线性代数                   ", "\t综合教学楼102", "\t周一 13:30-15:10");
        Course course2 = new Course(2, "\tJava应用开发技术", "\t综合教学楼305", "\t周四 15:30-17:10");
        Course course3 = new Course(3, "\t大学物理                   ", "\t主楼3402   ", "\t周四 10:00-11:40");
        
        // 实例化教师对象
        Teachers teacher1 = new Teachers(1, "王老师", "女", course1);
        Teachers teacher2 = new Teachers(2, "张老师", "男", course2);
        Teachers teacher3 = new Teachers(3, "谭老师", "男", course3);
        
        // 实例化学生对象
        Students student1 = new Students(1, "小明", "男");
        Students student2 = new Students(2, "小红", "女");
        
        //建立棵程和教师的联系
        course1.gteacher(teacher1);
        course2.gteacher(teacher2);
        course3.gteacher(teacher3);
        
        
        // 学生选课
        student1.selectCourse(course1);
        student2.selectCourse(course2);
        
        // 打印学生课表信息
        System.out.println("学生课表信息：");
        System.out.println("\t编号\t课程名称\t\t\t上课地点\t\t\t时间\t\t授课教师");
        printCourseInfo(student1);
        printCourseInfo(student2);
        
        // 学生退课
        student1.dropCourse();
        student2.dropCourse();
        
        //学生再选课
        student1.selectCourse(course2);
        student2.selectCourse(course3);
        
        // 打印学生课表信息
        System.out.println("\n学生退选课后的课表信息：");
        System.out.println("\t编号\t课程名称\t\t\t上课地点\t\t\t时间\t\t授课教师");
        printCourseInfo(student1);
        printCourseInfo(student2);
        
        // 保存选课结果到Excel文件
        saveToExcel("选课结果.xls", student1, student2);
    
    }
	public static void printCourseInfo(Students student) {
        if (student.getCourse() != null) {
            System.out.println(student.getnum() + "\t" + student.getCourse().getCourseName() + "\t" +
                    student.getCourse().getLocation() + "\t" + student.getCourse().getTime() +
                    "\t" + student.getCourse().teacher.name);
        } else {
            System.out.println(student.getnum() + "\t未选课\t-\t-\t-");
        }
    }
	public static void saveToExcel(String fileName, Students... students) {
		Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("选课结果");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("编号");
        cell = row.createCell(1);
        cell.setCellValue("姓名");
        cell = row.createCell(2);
        cell.setCellValue("性别");
        cell = row.createCell(3);
        cell.setCellValue("课程名称");
        cell = row.createCell(4);
        cell.setCellValue("上课地点");
        cell = row.createCell(5);
        cell.setCellValue("时间");
        cell = row.createCell(6);
        cell.setCellValue("授课教师");
        
        int rowNum = 1;
        for (Students student : students) {
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue(student.num);
            cell = row.createCell(1);
            cell.setCellValue(student.name);
            cell = row.createCell(2);
            cell.setCellValue(student.gender);
            if (student.getCourse() != null) {
                cell = row.createCell(3);
                cell.setCellValue(student.getCourse().courseName);
                cell = row.createCell(4);
                cell.setCellValue(student.getCourse().location);
                cell = row.createCell(5);
                cell.setCellValue(student.getCourse().time);
                cell = row.createCell(6);
                cell.setCellValue(student.getCourse().teacher.name);
            } else {
                cell = row.createCell(3);
                cell.setCellValue("未选课");
            }
        }
        
        try {
            FileOutputStream outputStream = new FileOutputStream(fileName);
            workbook.write(outputStream);
            outputStream.close();
            System.out.println("选课结果已保存到Excel文件：" + fileName);
        } catch (IOException e) {
            e.printStackTrace();
            }
	}
	
}
    
    
