package wtabae;
/*
 * (c) 2016, Aleksey Eremin
 * 
 * �������������� rtf ������ � ����������� � �������� ����������
 */

import java.io.File;

public class WordDirRtf extends myParser {
	
	@Override
	boolean parse(String directoryFileName, String maskFileName) {
		int i,n;
		String stri;
		boolean result=false;
		String DirFi;
		maskFileName=maskFileName+".*rtf$"; // ���� ������ RTF �����
		// �������� ���������� - ���� ��� � �������� �� �����������?
		File mdir=new File(directoryFileName);
		if(!mdir.exists() || !mdir.isDirectory()) {
			Log("?-ERROR-Not found directory:" + directoryFileName);
			return false;
		}
		// ������� ���� ����������,��� ����� �����
		DirFi=mdir.getAbsolutePath() + System.getProperty("file.separator");  // ������� �����������
		// ������� ������ docx ������, ��� ���� ��������� �������
		// �� ����� ���� ����� ���� 1, �� ��� ��� ��� �������� ��� Regexp ���������,
		// �� ����� ���� �����, � �� ���� �� ����������
		String[] list=mdir.list(new MyFilter(maskFileName));
		n=list.length;
		for(i=0; i<n; i++) {
			stri=DirFi + list[i]; // ��� RTF ����� � ��������
			System.out.println(stri);
			// ������� ����������� ������ �� ��������� �������
			WordTableRTF wt=new WordTableRTF();
			if(wt.parse(stri, stri)) {
				result=true;
			}
			//
		} // end for
		return result;
	} // end parse()	

} // end class
