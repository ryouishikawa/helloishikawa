import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

public class zentest{

  public static void main(String args[]) {
  Scanner scanner = new Scanner(System.in);
  System.out.print("yyyy/mm/dd形式で日付を入力してください。：");
  String date = scanner.next();	
  String wordhi = zengetu(date);
  System.out.println(wordhi);

}

public static String zengetu(String hiduke){
        String nowday = hiduke;

        //１月前を取得
        //hiduke を分解。
        String hidukenen = hiduke.substring(0,4);
        String hiduketuki = hiduke.substring(5,7);
        String hidukehi = hiduke.substring(8,10);
        String hiduketuki2 = hiduketuki;
        //前年算出
        int _nenwk = Integer.parseInt(hidukenen);
        int _nenwk2 = _nenwk - 1;

        //前月算出
        int _tukiwk = Integer.parseInt(hiduketuki);
        int _tukiwk2 = _tukiwk - 1;

        //前日算出
        int _hiwk = Integer.parseInt(hidukehi);

        String parahiduke = "";
        //前月の値格納ｗｋ
        String hidukezentuki = "";
        String hiduketukiwk = "";

        //前月桁落ち防止用ｗｋ
        String hidukezentukiwk="";
        //日桁落ち防止用ｗｋ
        String hidukehiwk="";
        //日の判定用
        int hidukehiint = 0;

        parahiduke = hiduke;

        //前日同日取得

        //前月同日用wk
        int nen0 = 0;
        int tuki0 = 0;
        int hi0 = 0;
        String snen0 = "";
        String stuki0 = "";
        String shi0 = "";

        // 1月対応
        if (_tukiwk2==0){
           _tukiwk2 = 12;
        }
        else{
        }        
       
        // 1月対応(前年を取得する）
        if (_tukiwk2==12){
           nen0 = _nenwk2;
        }
        else{
           nen0 = _nenwk;
        }        
        snen0 = String.valueOf(nen0);


        // 桁落ち対応（月）
        tuki0 = _tukiwk2;

        if (tuki0<10){
            stuki0 = "0" + String.valueOf(tuki0);
        }
        else{
            stuki0 = String.valueOf(tuki0);
        }        


        // 桁落ち対応（日）
        hi0 = _hiwk;


        // 月末対応（日）
        if ((hi0>30)&&(tuki0==1||tuki0==3||tuki0==5||tuki0==7||tuki0==8||tuki0==10||tuki0==12)){
		hi0 = 31;		
        }else{
        }        

        if ((hi0>30)&&(tuki0==4||tuki0==6||tuki0==9||tuki0==11)){
		hi0 = 30;		
        }else{
        }        
      
        if ((hi0>28)&&(tuki0==2)){
                if((nen0%4)==0){
		hi0 = 28;	
                }else{
		hi0 = 29;	
                }
        }else{
        }        

        if (hi0<10){
            shi0 = "0" + String.valueOf(hi0);
        }
        else{
            shi0 = String.valueOf(hi0);
        }        

        parahiduke = snen0 + "/" + stuki0 + "/" + shi0;

return parahiduke;

}

}