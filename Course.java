package XuanKe;

public class Course {
	 int id;
	 public String courseName;
	 public String location;
	 public String time;
	 public Teachers teacher;
	
	public Course(int id, String courseName, String location, String time) {
        this.id = id;
        this.courseName = courseName;
        this.location = location;
        this.time = time;
    }
	public String getCourseName() {
        return id+"  "+courseName;
    }
	public String getLocation() {
		// TODO Auto-generated method stub
		return location;
	}
	public String getTime() {
		// TODO Auto-generated method stub
		return time;
	}
	
	
	public void gteacher(Teachers t){
		this.teacher=t;}
	
	
	
}
