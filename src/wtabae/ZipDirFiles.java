package wtabae;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/*
 *  (C), 2016 ������� ������
 *  
 *  ������������ ZIP ������
 */

public class ZipDirFiles {
    String dirFiles;

    // ������� ������� �������
    ZipDirFiles(String inputDir)
    {
            dirFiles=inputDir;
    }

    // ������������ ����� ������ � ����������� ext � ��������� �������� � �����
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
        // ������� ���� ����������,��� ����� �����
        DirFi=mdir.getAbsolutePath() + System.getProperty("file.separator");  // ������� �����������
        // ������� ������ ������, ���� ���� ���������
        String[] list=mdir.list(new MyFilter(".*"+ext+"$"));
        n=list.length;
        // ���������
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
            stri = list[i]; // ��� RTF ����� � ��������
            // ����������
            ZipEntry ze=new ZipEntry(stri); // ������� ������ ��� ������ (��������� � ��������)
            try {
                zos.putNextEntry(ze);       // ���������� ������ � �����
                // ��������� ���� � ������� ������ ������
                byte[] readBuffer=new byte[2048];
                int bytesRead=0;
                // ��������� ������� ����
                FileInputStream fis= new FileInputStream(DirFi+stri);
                cnt=0;
                while((bytesRead=fis.read(readBuffer)) != -1) {
                    cnt += bytesRead;
                    zos.write(readBuffer, 0, bytesRead);
                } // end while
                fis.close();    // ��������� ������� ����
                zos.closeEntry(); // ��������� ������ � ������
                result++;
                System.out.println(result + ") " + stri);
            } catch (IOException ex) {
                ex.printStackTrace();
                System.out.println("?-WARNING-can't add to zip file: " + stri);
            }
            //
        } // end for
        // ��������� ��� �����
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
    // ������������ ����� ������ � ����������� ext � ��������� �������� � ����� 
    // � ������� ������� pkzipc (������ ���� � C:\Windows)
    public int makeZipPkzip(String ext, String zipFile)
    {
    	String DirFi, stri, zipname;
    	String strcmd, str;
    	File mdir=new File(dirFiles);
        if(!mdir.exists() || !mdir.isDirectory()) {
            System.out.println("?-ERROR-Not found directory:" + dirFiles);
            return 0;
        }
        // ������� ���� ����������,��� ����� �����
        DirFi=mdir.getAbsolutePath() + System.getProperty("file.separator");  // ������� �����������
        ext="*." + ext;	// *.rtf
        // ������ ��� ������
        zipname=DirFi+zipFile+".zip";
        //
        strcmd="pkzipc -add -silent " + zipname + " " + DirFi + ext;
        str=RunWord.execOS(strcmd);
        // ���������� �����
        System.out.println(strcmd);
        System.out.println(str);
        //
    	return 1;
    }
    

}
