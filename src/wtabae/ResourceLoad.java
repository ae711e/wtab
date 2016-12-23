/*
 * Copyright (c) 2016. Aleksey Eremin
 * 23.12.16 9:37
 */

package wtabae;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

/**
 * Created by ae on 23.12.2016.
 * ��������� �������
 */

public class ResourceLoad {
  
  // ��������� ������ �� ����� ������� (��-��������� ��������� Cp1251)
  public String getText(String name)
  {
    return getText(name, "Cp1251");
  }
  
  // ��������� ������ �� ����� �������
  public String getText(String name, String code_page)
  {
    StringBuilder sb = new StringBuilder();
    try {
      InputStream is = this.getClass().getResourceAsStream(name);  // ��� �������
      BufferedReader br = new BufferedReader(new InputStreamReader(is, code_page));
      String line;
      while ((line = br.readLine()) !=null) {
        sb.append(line);  sb.append("\n");
      }
    } catch (IOException ex) {
      ex.printStackTrace();
      // * @author Eugene Matyushkin aka Skipy
      // @since 03.08.12
      //StringWriter sw = new StringWriter();
      //PrintWriter pw = new PrintWriter(sw);
      //ex.printStackTrace(pw);
      //pw.flush();
      //pw.close();
      //sb.append("Error while loading text: ").append("\n\n");
      //sb.append(sw.getBuffer().toString());
      //
    }
    return sb.toString();
  }
  
}
