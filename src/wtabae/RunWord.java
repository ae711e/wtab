package wtabae;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;

/*
 *  ��������� rtf � docx � ������� � ������� VBS Word
 */

public class RunWord {
	// ����� VBS ��������� ��� ������� Word � �������������� ����� RTF - DOCX
	final static String bat_vbs="' (C) 2016 Aleksey Eremin (vbs) \r\n"
			+ "Dim WA, WD, Args \r\n"
			+ "Set Args = WScript.Arguments \r\n"
			+ "Set WA = CreateObject(\"Word.Application\") \r\n"
			+ "Set WD = WA.Documents.Open(Args(0)) \r\n"
			+ "WD.SaveAs Args(1), %d \r\n"
			+ "WD.Close \r\n"
			+ "Set WD = Nothing \r\n"
			+ "WA.Quit \r\n"
			+ "Set WA = Nothing \r\n";
		

	// ������������� RTF ���� � DOCX ����
	// ����� ������� ���������� ����� � VBS ��������
	public static boolean rtf2docx(String fileRtf, String fileDocx) 
	{
		double d=Math.random()*1000;
		int irnd=(int)d+1;	// ��������� ����� 1-1000
		return vbsFileExec(16, "ard" + irnd + ".vbs", fileRtf, fileDocx); // rtf2docx.vbs				
	} // end rtf2docx()
	
	// ������������� DOCX ���� � RTF ����
	// ����� ������� ���������� ����� � VBS ��������
	public static boolean docx2rtf(String fileDocx, String fileRtf) 
	{
		double d=Math.random()*1000;
		int irnd=(int)d+1;	// ��������� ����� 1-1000
		return vbsFileExec(6, "adr" + irnd + ".vbs", fileDocx, fileRtf); // docx2rtf.vbs
	} // end docx2rtf()
		
    // ������ �� ���������� Word ����� ������� ���������� ����� � VBS ��������
	// ������������ ��������� ���� VBS, ��������� ��� � ������� � �������� ������
	// ��� ������ ���������� 1
	private static boolean vbsFileExec(int cod, String fileNameVbs, String file1, String file2) 
	{
		String str, fileVBS, fileCMD;
		 // ���������� ��������� ���� VBS
		 File fvbs=new File(tmpDir(), fileNameVbs);		 
		 fileVBS=fvbs.getPath();
		 str=String.format(bat_vbs, cod);
		 try {
			 PrintWriter out = new PrintWriter(fileVBS);
			 out.write(str);
			 out.close();
		 } catch(IOException ex) {
			 ex.printStackTrace();
			 System.out.println("?-ERROR-can't create vbs file-"+fileVBS);
			 return false;
		 }
		 // ������ �������
		 str="\"" + fileVBS + "\" \"" + file1 + "\" \"" + file2 + "\"";
		 // ����� ���������� ������ file2
		 File fs=new File(file2);
		 fs.delete();		// ������ �������� ����
		 //
		 double d=Math.random()*1000;
		 int irnd=(int)d+1;	// ��������� ����� 1-1000
		 // �������� ��������� ����
		 File fcmd=new File(tmpDir(), "cae" + irnd + ".bat");
		 fileCMD=fcmd.getPath();
		 try {
			 PrintWriter out = new PrintWriter(fileCMD);
			 out.write(str);
			 out.close();
		 } catch(IOException ex) {
			 ex.printStackTrace();
			 System.out.println("?-ERROR-can't create cmd file-"+fileCMD);
			 return false;
		 }
		 //		 
		 str=execOS("CMD /C " + fileCMD);
		 // ���������� �����
		 /////System.out.println(str);
		 //
		 fvbs.deleteOnExit(); // ������� �� ���������� ��������� ���� VBS
		 fcmd.deleteOnExit(); // ������� ��������� ���� �� ���������� ��������� CMD
		 // ���� �������� ���� �������� - ������ OK
		 return fs.exists();		 
	} // end vbsFileExec()

	// ��������� ������� ��
	public static String execOS(String command)
	{
		String result="";
		Runtime r = Runtime.getRuntime();
		try {
			//Process p = r.exec("cmd /c "+command);
			Process p = r.exec(command);
			p.waitFor();
			BufferedReader b = new BufferedReader(new InputStreamReader(p.getInputStream()));
			String line;
			//
			while ((line = b.readLine()) != null) {			  
			  result += line;
			}
			b.close();
		}
		catch(Exception ex) {
			ex.printStackTrace();
			return "?-ERROR-" + command;
		}				
		return result;
	}
	
	// ������ ��������� ������� (����������� �������� ������)
	public static String tmpDir() 
	{
		return System.getProperty("java.io.tmpdir");
	}

} // end class


/*
����� �������� 
'rtf2docx.vbs
Dim WA, WD, Args
Set Args = WScript.Arguments
Set WA = CreateObject("Word.Application")
Set WD = WA.Documents.Open(Args(0))
WD.SaveAs Args(1), 16
WD.Close
Set WD = Nothing
WA.Quit
Set WA = Nothing

'docx2rtf
Dim WA, WD, Args
Set Args = WScript.Arguments
Set WA = CreateObject("Word.Application")
Set WD = WA.Documents.Open(Args(0))
WD.SaveAs Args(1), 6
WD.Close
Set WD = Nothing
WA.Quit
Set WA = Nothing


*/
