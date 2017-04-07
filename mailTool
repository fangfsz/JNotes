package mailTool;
import java.io.File;
import java.util.Calendar;
import java.util.Date;
import java.util.Enumeration;
import java.util.GregorianCalendar;
import java.util.Vector;

import lotus.domino.Database;
import lotus.domino.DateRange;
import lotus.domino.DbDirectory;
import lotus.domino.Document;
import lotus.domino.DocumentCollection;
import lotus.domino.EmbeddedObject;
import lotus.domino.Item;
import lotus.domino.NotesException;
import lotus.domino.NotesFactory;
import lotus.domino.NotesThread;
import lotus.domino.RichTextItem;
import lotus.domino.Session;
import lotus.domino.View;
public class MailCollect extends NotesThread{

	public static void main(String[] args) {
		// TODO Auto-generated method stub
			MailCollect t = new MailCollect();		
			t.start();
	}
	public void runNotes() {
		openDBGetFilesByDate(2017, 3, 20);
	}

	
	public void openDBGetFilesByDate(int y, int m, int d) {
		Calendar lc = new GregorianCalendar();
		if (y > lc.get(Calendar.YEAR) || y < 2014) {
			System.out.println("Year must be >=2015 and <= current year.");
		}
		if (m < 0 || m > 12) {
			System.out.println("Monther must be between 1~12.");
		}
		if (d < 0 || d > lc.getActualMaximum(Calendar.DAY_OF_MONTH)) {
			System.out.println("Days in Month must be valid.");
		}
		if (new Calendar.Builder().setDate(y, m - 1, d).build() == null) {
			System.out.println("Invalid Date format!!!");
		}
		Session s;
		try {
			Calendar lc1 = new GregorianCalendar();
			lc1.set(Calendar.YEAR, y);
			lc1.set(Calendar.MONTH, m - 1);
			lc1.set(Calendar.DAY_OF_MONTH, d);
			//当天上午7点
			lc1.set(Calendar.HOUR_OF_DAY, 7);
			lc1.set(Calendar.MINUTE, 0);
			lc1.set(Calendar.SECOND, 0);
			Date d1 = lc1.getTime();
			System.out.println(d1.toString());
			Calendar lc2 = new GregorianCalendar();
			lc2.set(Calendar.YEAR, y);
			lc2.set(Calendar.MONTH, m - 1);
			lc2.set(Calendar.DAY_OF_MONTH, d);
			//当天下午16:30
			lc2.set(Calendar.HOUR_OF_DAY, 20);
			lc2.set(Calendar.MINUTE, 30);
			lc2.set(Calendar.SECOND, 0);
			Date d2 = lc2.getTime();
			System.out.println(d2.toString());
			s = NotesFactory.createSession();
			//主要是获取当天指定时间段内的邮件
			DateRange dr = s.createDateRange(s.createDateTime(lc1.getTime()), s.createDateTime(lc2.getTime()));
			DbDirectory Dir = s.getDbDirectory(null);
			Database db = Dir.openMailDatabase();
			String fn = "SMI Offline Alert";
			if (db == null)
				System.out.println("Db is not opened.");
			System.out.println("Db ( " + db.getFileName() + " ) is opened now. ");
			DocumentCollection dc = db.getAllDocuments();
			System.out.println("Mail database : " + db.getTitle() + "'s size is " + ((int) (db.getSize() / 1024 / 1024))
					+ "MB and has " + dc.getCount() + " documents");
			View folder = db.getView(fn);
			System.out.println("The folder(" + fn + ") has " + folder.getEntryCount() + " docs.");
			System.out.println("------------Start saving the fileattachment----------");
			Document doc = folder.getFirstDocument();
			Document tmpdoc;
			int c = 0;
			System.out.println("Preparing dowloading the SMI alerts for " + dr.getText());
			while (doc != null && folder.getEntryCount() != 0) {				
				   
						Item ibody = doc.getFirstItem("Body");
						if (ibody instanceof RichTextItem) {
							RichTextItem rtitem = (RichTextItem) ibody;
							// System.out.println("RichText Name: "+
							// rtitem.getName()+"\n RichText Type: "+ rtitem.getType());
							Vector<?> v = rtitem.getEmbeddedObjects();
							Enumeration<?> e = v.elements();
							 Date dt1= doc.getLastModified().toJavaDate();
//						    System.out.println(dt1.toString());
						    if((dt1.before(d2) && dt1.after(d1))){
						    
								while (e.hasMoreElements()) {
									EmbeddedObject eo = (EmbeddedObject) e.nextElement();
									String regx = "^(?:(?!NNJA)[A-Za-z])+[A-Za-z-]?[A-Za-z]+[A-Za-z-]?[A-Za-z]+_(?:UX|WIN)_\\d{8}\\.out$";
									if (eo.getType() == EmbeddedObject.EMBED_ATTACHMENT && eo.getName().matches(regx)) {
										// Create the directory firstly.
										String dirName="c:\\testMail\\";
										File file= new File(dirName);
										file.mkdirs();
										eo.extractFile(dirName + eo.getSource());
										++c;
										System.out.println(c + "-->" + eo.getName());
										// default override the file if same filename.
									}
								}

						    }
							tmpdoc = folder.getNextDocument(doc);
							doc.recycle();
							doc = tmpdoc;
						

						}
				
				    
			}

			System.out.println("------------Total   " + c + "   docs got!----------");
		} catch (NotesException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}


}
