/*
 * Copyright (c) 2016. Aleksey Eremin
 * 23.12.16 9:27
 */

package wtabae;

import java.io.File;

public class Wtab {

	public static void main(String[] args) {
          // TODO Auto-generated method stub
          //
          try {
              // изменения внес дома
              // изменил отступы (на работе)
              String a1="C:\\TMP\\_REV\\monitoring_protocol_";
              //
            if (args.length>0) {
              a1=args[0];	// директория и маска имени файла
            } else {
              ResourceLoad rl = new ResourceLoad();  // подготовим загрузку ресурсов
              String hello=rl.getText("/res/hello.txt");
              //System.out.println(hello);
              System.out.printf(hello, a1);
            }
            //
            File nf=new File(a1);
            a1=nf.getParent();
            String a2=nf.getName();
            // начнем обработку
            myParser pars=new WordDirRtf();//обработка RTF файлов
            
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

