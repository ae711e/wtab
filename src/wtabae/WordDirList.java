package wtabae;
// (C) 2016 Aleksey Eremin

import java.io.File;

// проходит по файлам в каталоге (directoryFileName),
// анализируютс€ файла RTF с маской имени (docxFileName)
public class WordDirList extends myParser
{

	@Override
	boolean parse(String directoryFileName, String docxFileName) {
		// TODO Auto-generated method stub
		int i,n;
		String stri;
		String DirFi;
		docxFileName=docxFileName+".*docx$"; // ищем только docx файлы
		// проверим директорию - есть она и €вл€етс€ ли директорией?
		File mdir=new File(directoryFileName);
		if(!mdir.exists() || !mdir.isDirectory()) {
			Log("?-ERROR-Not found directory:" + directoryFileName);
			return false;
		}
		// получим путь директории,где файлы лежат
		DirFi=mdir.getAbsolutePath() + System.getProperty("file.separator");  // добавим разделитель
		// получим список docx файлов, где надо поправить таблицу
		// на самом деле такой файл 1, но так как им€ задаетс€ как Regexp выражение,
		// их может быть много, и мы всех их обработаем
		String[] list=mdir.list(new MyFilter(docxFileName));
		n=list.length;
		for(i=0; i<n; i++) {
			stri=DirFi + list[i]; // им€ RTF файла в каталоге
			System.out.println(stri);
			// заведем родственный объект по обработке таблицы
			WordTableDOCX wt=new WordTableDOCX();
			wt.parse(stri, stri);
			//
		} // end for
		return true;
	}

	//
}
