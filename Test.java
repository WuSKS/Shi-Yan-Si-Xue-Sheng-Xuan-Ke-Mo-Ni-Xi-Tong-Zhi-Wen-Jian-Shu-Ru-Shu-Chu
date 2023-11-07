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
        // ʵ�����γ̶���
        Course course1 = new Course(1, "\t���Դ���                   ", "\t�ۺϽ�ѧ¥102", "\t��һ 13:30-15:10");
        Course course2 = new Course(2, "\tJavaӦ�ÿ�������", "\t�ۺϽ�ѧ¥305", "\t���� 15:30-17:10");
        Course course3 = new Course(3, "\t��ѧ����                   ", "\t��¥3402   ", "\t���� 10:00-11:40");
        
        // ʵ������ʦ����
        Teachers teacher1 = new Teachers(1, "����ʦ", "Ů", course1);
        Teachers teacher2 = new Teachers(2, "����ʦ", "��", course2);
        Teachers teacher3 = new Teachers(3, "̷��ʦ", "��", course3);
        
        // ʵ����ѧ������
        Students student1 = new Students(1, "С��", "��");
        Students student2 = new Students(2, "С��", "Ů");
        
        //�����ó̺ͽ�ʦ����ϵ
        course1.gteacher(teacher1);
        course2.gteacher(teacher2);
        course3.gteacher(teacher3);
        
        
        // ѧ��ѡ��
        student1.selectCourse(course1);
        student2.selectCourse(course2);
        
        // ��ӡѧ���α���Ϣ
        System.out.println("ѧ���α���Ϣ��");
        System.out.println("\t���\t�γ�����\t\t\t�Ͽεص�\t\t\tʱ��\t\t�ڿν�ʦ");
        printCourseInfo(student1);
        printCourseInfo(student2);
        
        // ѧ���˿�
        student1.dropCourse();
        student2.dropCourse();
        
        //ѧ����ѡ��
        student1.selectCourse(course2);
        student2.selectCourse(course3);
        
        // ��ӡѧ���α���Ϣ
        System.out.println("\nѧ����ѡ�κ�Ŀα���Ϣ��");
        System.out.println("\t���\t�γ�����\t\t\t�Ͽεص�\t\t\tʱ��\t\t�ڿν�ʦ");
        printCourseInfo(student1);
        printCourseInfo(student2);
        
        // ����ѡ�ν����Excel�ļ�
        saveToExcel("ѡ�ν��.xls", student1, student2);
    
    }
	public static void printCourseInfo(Students student) {
        if (student.getCourse() != null) {
            System.out.println(student.getnum() + "\t" + student.getCourse().getCourseName() + "\t" +
                    student.getCourse().getLocation() + "\t" + student.getCourse().getTime() +
                    "\t" + student.getCourse().teacher.name);
        } else {
            System.out.println(student.getnum() + "\tδѡ��\t-\t-\t-");
        }
    }
	public static void saveToExcel(String fileName, Students... students) {
		Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("ѡ�ν��");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("���");
        cell = row.createCell(1);
        cell.setCellValue("����");
        cell = row.createCell(2);
        cell.setCellValue("�Ա�");
        cell = row.createCell(3);
        cell.setCellValue("�γ�����");
        cell = row.createCell(4);
        cell.setCellValue("�Ͽεص�");
        cell = row.createCell(5);
        cell.setCellValue("ʱ��");
        cell = row.createCell(6);
        cell.setCellValue("�ڿν�ʦ");
        
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
                cell.setCellValue("δѡ��");
            }
        }
        
        try {
            FileOutputStream outputStream = new FileOutputStream(fileName);
            workbook.write(outputStream);
            outputStream.close();
            System.out.println("ѡ�ν���ѱ��浽Excel�ļ���" + fileName);
        } catch (IOException e) {
            e.printStackTrace();
            }
	}
	
}
    
    
