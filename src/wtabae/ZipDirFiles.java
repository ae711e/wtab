package wtabae;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/*
 *  (C), 2016 Алексей Еремин
 *  
 *  формирование ZIP архива
 */

public class ZipDirFiles {
    String dirFiles;

    // указать рабочик каталог
    ZipDirFiles(String inputDir)
    {
            dirFiles=inputDir;
    }

    // сформировать архив файлов с расширением ext в известном каталоге в архив
    public int makeExt2Zip(String ext, String zipFile)
    {
        String DirFi, stri, zipname;
        int result=0;
        int i, n, cnt;
        File mdir=new File(dirFiles);
        if(!mdir.exists() || !mdir.isDirectory()) {
            System.out.println("?-ERROR-Not found directory:" + dirFiles);
            return 0;
        }
        // получим путь директории,где файлы лежат
        DirFi=mdir.getAbsolutePath() + System.getProperty("file.separator");  // добавим разделитель
        // получим список файлов, кого надо упаковать
        String[] list=mdir.list(new MyFilter(".*"+ext+"$"));
        n=list.length;
        // архиватор
        zipname=DirFi+zipFile+".zip";
        ZipOutputStream zos=null;
        FileOutputStream out=null;
        try {
            out=new FileOutputStream(zipname);
            zos= new ZipOutputStream(out);
        } catch (IOException ex) {
            ex.printStackTrace();
            System.out.println("?-ERROR-can't create zip file: " + zipname);
            return 0;
        }
        //
        System.out.println("Dir list file to ZIP:");
        for(i=0; i<n; i++) {
            stri = list[i]; // имя RTF файла в каталоге
            // архивируем
            ZipEntry ze=new ZipEntry(stri); // создаем запись для архива (указываем её название)
            try {
                zos.putNextEntry(ze);       // вкладываем запись в архив
                // считываем файл в текущую запись архива
                byte[] readBuffer=new byte[2048];
                int bytesRead=0;
                // открываем входной файл
                FileInputStream fis= new FileInputStream(DirFi+stri);
                cnt=0;
                while((bytesRead=fis.read(readBuffer)) != -1) {
                    cnt += bytesRead;
                    zos.write(readBuffer, 0, bytesRead);
                } // end while
                fis.close();    // закрываем входной файл
                zos.closeEntry(); // закрываем запись в архиве
                result++;
                System.out.println(result + ") " + stri);
            } catch (IOException ex) {
                ex.printStackTrace();
                System.out.println("?-WARNING-can't add to zip file: " + stri);
            }
            //
        } // end for
        // закрываем все нафиг
        try {
            zos.close();
            out.close();
        } catch (IOException ex) {
            ex.printStackTrace();
            System.out.println("?-ERROR-can't close zip: " + zipname);
        }
        //
        return result;
    }
    
    // pkzipc -add " & DirFi & "_a.zip " & DirFi & "monitoring_*.rtf"
    // сформировать архив файлов с расширением ext в известном каталоге в архив 
    // с помощью утилиты pkzipc (должна быть в C:\Windows)
    public int makeZipPkzip(String ext, String zipFile)
    {
    	String DirFi, stri, zipname;
    	String strcmd, str;
    	File mdir=new File(dirFiles);
        if(!mdir.exists() || !mdir.isDirectory()) {
            System.out.println("?-ERROR-Not found directory:" + dirFiles);
            return 0;
        }
        // получим путь директории,где файлы лежат
        DirFi=mdir.getAbsolutePath() + System.getProperty("file.separator");  // добавим разделитель
        ext="*." + ext;	// *.rtf
        // полное имя архива
        zipname=DirFi+zipFile+".zip";
        //
        strcmd="pkzipc -add -silent " + zipname + " " + DirFi + ext;
        str=RunWord.execOS(strcmd);
        // отладочный вывод
        System.out.println(strcmd);
        System.out.println(str);
        //
    	return 1;
    }
    

}
