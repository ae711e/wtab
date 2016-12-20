package wtabae;
/*
 * (c) 2016, Aleksey Eremin
 * 
 * Преобразование rtf файлов с протоколами в заданной директории
 */

import java.io.File;

public class WordDirRtf extends myParser {
	
	@Override
	boolean parse(String directoryFileName, String maskFileName) {
		int i,n;
		String stri;
		boolean result=false;
		String DirFi;
		maskFileName=maskFileName+".*rtf$"; // ищем только RTF файлы
		// проверим директорию - есть она и является ли директорией?
		File mdir=new File(directoryFileName);
		if(!mdir.exists() || !mdir.isDirectory()) {
			Log("?-ERROR-Not found directory:" + directoryFileName);
			return false;
		}
		// получим путь директории,где файлы лежат
		DirFi=mdir.getAbsolutePath() + System.getProperty("file.separator");  // добавим разделитель
		// получим список docx файлов, где надо поправить таблицу
		// на самом деле такой файл 1, но так как имя задается как Regexp выражение,
		// их может быть много, и мы всех их обработаем
		String[] list=mdir.list(new MyFilter(maskFileName));
		n=list.length;
		for(i=0; i<n; i++) {
			stri=DirFi + list[i]; // имя RTF файла в каталоге
			System.out.println(stri);
			// заведем родственный объект по обработке таблицы
			WordTableRTF wt=new WordTableRTF();
			if(wt.parse(stri, stri)) {
				result=true;
			}
			//
		} // end for
		return result;
	} // end parse()	

} // end class
