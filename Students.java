package XuanKe;

public class Students {
	 public int num;
	 public String name;
	 public String gender;
	 public Course course;
	
	
	
	
	
	
	public Students(int num,String name,String gender){
		this.num = num;
		this.name= name;
		this.gender= gender;
		
	}
	public void selectCourse(Course course) {
        this.course = course;
    }
	
	
    
    public void dropCourse() {
        this.course = null;
    }
    
    public Course getCourse() {
        return course;
    }
    
	public String getnum() {
		// TODO Auto-generated method stub
		return name;
	}
	
	
    

	
}
