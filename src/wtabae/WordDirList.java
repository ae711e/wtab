package wtabae;
// (C) 2016 Aleksey Eremin

import java.io.File;

// �������� �� ������ � �������� (directoryFileName),
// ������������� ����� RTF � ������ ����� (docxFileName)
public class WordDirList extends myParser
{

	@Override
	boolean parse(String directoryFileName, String docxFileName) {
		// TODO Auto-generated method stub
		int i,n;
		String stri;
		String DirFi;
		docxFileName=docxFileName+".*docx$"; // ���� ������ docx �����
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
		String[] list=mdir.list(new MyFilter(docxFileName));
		n=list.length;
		for(i=0; i<n; i++) {
			stri=DirFi + list[i]; // ��� RTF ����� � ��������
			System.out.println(stri);
			// ������� ����������� ������ �� ��������� �������
			WordTableDOCX wt=new WordTableDOCX();
			wt.parse(stri, stri);
			//
		} // end for
		return true;
	}

	//
}
