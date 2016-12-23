/*
 * Copyright (c) 2016. Aleksey Eremin
 *
 * Created by ae on 06.07.2016.
 *
 * Обработать таблицы в документе RTF
 * Для этого RTF переведем в DOCX
 * Обработаем DOCX
 *
 */

package wtabae;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

public class WordTableRTF extends WordTableDOCX {

    @Override
    boolean parse(String inputFileName, String outputFileName) {
        // return super.parse(inputFileName, outputFileName);
    	// сначала сделаем копию входного файла
    	boolean result=false;
    	String pfile;	// имя промежуточного файла
    	//
    	pfile=inputFileName + ".docx";
    	// преобразуем входной файл rtf в docx
    	if (RunWord.rtf2docx(inputFileName, pfile)) {
    		// переименуем входной файл в .bak
    		File finp=new File(inputFileName);
    		File fbak=new File(inputFileName+".bak");
    		if(!finp.renameTo(fbak)) {
    			Log("?-WARNING-can't rename file: " + inputFileName);
    		}
    		if(super.parse(pfile, pfile)) {
	    		if(RunWord.docx2rtf(pfile, inputFileName)) {
	    			fbak.delete();
	    			File f=new File(pfile);
	    			f.delete();
	    			result=true;
	    		}else {
	    			Log("?-ERROR-can't convert to RTF " + pfile);
	    			fbak.renameTo(finp);
	    			Log("?-WARNING-restore input file: "+ inputFileName);
				}
    		}else {
    		    File f=new File(pfile);
    		    f.delete();
    		    fbak.renameTo(finp);
    		}
    	}
    	return result;
    } // end parse()
    
    
    // простой и удобный метод копирования файла в Java 7
    public static boolean copyFile2File(File source, File dest) 
    {
    	try {
        	Files.copy(source.toPath(), dest.toPath());
        	return true;
        } catch (IOException ex){
        	ex.printStackTrace();
        }
        return false;
    } // end copyFile2File()
    
    // простой и удобный метод копирования файла в Java 7
    public static void ex_copyFile2File(File source, File dest) throws IOException 
    {
    	Files.copy(source.toPath(), dest.toPath());        	
    } // end copyFile2File()
    
}
