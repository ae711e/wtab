package wtabae;

import java.io.File;

public class Wtab {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		//
          // ������� � 17:32
          
		try {
		    // str=WordParser.parsed("c:/tmp/a97.doc");
		    // str=WordParser.parsex("c:/tmp/_rev/a.docx");
		    // myParser pars=new WordTableDOCX();
		    // myParser pars=new WordTable2CSV();
			//
		    String a1="C:\\TMP\\_REV\\monitoring_protocol_";		    
		    //
		    if (args.length>0) {
		    	a1=args[0];	// ���������� � ����� ����� �����
		    } else {
                        System.out.println("(C) 2016 ������� ������");
                        System.out.println("WTAB v.2.06 02.12.2016");
                        System.out.println("������� �� ������� ���������� ������ � ��������������� ������� .jpg");
                        System.out.println("c ������� Word ������������ ����������� RTF � DOCX � �������");
                        System.out.println("c ������� pkzipc ��������� ����� zip");
                        System.out.println(">java -jar wtab.jar Dir\\File");
                        System.out.println("Dir  - ����������, ��� ���������  ���� � �������� � ����� ����������");
                        System.out.println("File - ���� Word RTF � �������� ���������� (�������� ��� ������ ����� �����)");
                        System.out.println("�� ��������� - " + a1);
			System.out.println("� Dir ����������� ����� _a.zip � ������� ���� � ���������, ���� ���������� ���-�� ����������.");
			System.out.println("�� ��������� ������ � ������ ������ ���������� ���������.");
                        System.out.println(" ");
		    }
		    //
		    File nf=new File(a1);
		    a1=nf.getParent();
                    String a2=nf.getName();
		    // ������ ���������
		    myParser pars=new WordDirRtf();//������� ������
		    //
		    if(pars.parse(a1, a2)) {
                      // ����������
                      ZipDirFiles zd=new ZipDirFiles(a1);
                      //zd.makeExt2Zip("rtf", "_aaa"); // �� ������� �� ����� ����� ������
                      zd.makeZipPkzip("rtf", "_a"); // ���������� ������� pkzipc
		    }
		    //
		}
		catch (Exception ex) {
                    ex.printStackTrace();
		}
		//
		// System.out.println("END PROGRAMM"); 
	} // end of main()
	

}

