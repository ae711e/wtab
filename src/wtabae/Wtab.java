package wtabae;

import java.io.File;

public class Wtab {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		//
          // добавил в 17:32
          // добавил в 17:34
          // добавил в 17:35
          
		try {
		    // str=WordParser.parsed("c:/tmp/a97.doc");
		    // str=WordParser.parsex("c:/tmp/_rev/a.docx");
		    // myParser pars=new WordTableDOCX();
		    // myParser pars=new WordTable2CSV();
			//
		    String a1="C:\\TMP\\_REV\\monitoring_protocol_";		    
		    //
		    if (args.length>0) {
		    	a1=args[0];	// директория и маска имени файла
		    } else {
                        System.out.println("(C) 2016 Алексей Еремин");
                        System.out.println("WTAB v.2.06 02.12.2016");
                        System.out.println("Удаляет из таблицы скриншотов строки с несуществующими файлами .jpg");
                        System.out.println("c помощью Word промежуточно преобразует RTF в DOCX и обратно");
                        System.out.println("c помощью pkzipc формирует архив zip");
                        System.out.println(">java -jar wtab.jar Dir\\File");
                        System.out.println("Dir  - директория, где находятся  файл с таблицей и файлы скриншотов");
                        System.out.println("File - файл Word RTF с таблицей скриншотов (задается как начало имени файла)");
                        System.out.println("По умолчанию - " + a1);
			System.out.println("В Dir формируется архив _a.zip с файлами акта и протокола, если изменилось кол-во скриншотов.");
			System.out.println("По окончанию работы в буфере обмена количество нарушений.");
                        System.out.println(" ");
		    }
		    //
		    File nf=new File(a1);
		    a1=nf.getParent();
                    String a2=nf.getName();
		    // начнем обработку
		    myParser pars=new WordDirRtf();//РАБОЧАЯ ВЕРСИЯ
		    //
		    if(pars.parse(a1, a2)) {
                      // запаковать
                      ZipDirFiles zd=new ZipDirFiles(a1);
                      //zd.makeExt2Zip("rtf", "_aaa"); // АС Ревизор не берет такие архивы
                      zd.makeZipPkzip("rtf", "_a"); // использует утилиту pkzipc
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

