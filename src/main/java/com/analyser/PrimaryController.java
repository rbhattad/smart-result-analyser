package com.herewegoagain;
//
//import java.io.File;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.regex.*;
//import org.apache.pdfbox.pdmodel.PDDocument;
//import org.apache.pdfbox.text.PDFTextStripper;  
//import org.apache.poi.ss.usermodel.IndexedColors;  
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.ss.util.CellRangeAddress;  
//import org.apache.poi.ss.usermodel.*;  
//import org.apache.poi.xssf.usermodel.XSSFCellStyle;  
//class Converter{
//    File pdffile;
//    public static int subn[]=new int[22];
//    public static String sub[]=new String[22];
//     public static int startfactor,endfactor,sem1end,sem2end;
//     public static boolean sem2;
//     public static void gocheck(String student,int y)
//        {
//            //function to set starting point of marks to finish point
//             Pattern total=Pattern.compile("Tot%");
//             Matcher t=total.matcher(student);
//             String year="";
//          switch (y) {
//             case 2:
//                 year= "20";
//                 break;
//             case 3:
//                 year = "30";
//                 break;
//             case 4:
//                 year = "40";
//                 break;
//         }
//             Pattern subject=Pattern.compile(year+"4181") ;
//             Matcher subjectline=subject.matcher(student);
//             t.find();
//             subjectline.find();
//             int i=t.end();
//             while(student.charAt(i+1)!='\n')// for total
//             {
//                 i++;
//             }
//             int j=subjectline.end();
//             while(student.charAt(j+1)!='\n')  //for subjectline
//             {
//                 j++;
//             }
//            // System.out.println(student.substring(subjectline.start(), j));
//             int end=j;
//             boolean forend=true;
//             while(true)
//             {
//                 if(student.charAt(i)=='T' && student.charAt(i+1)=='o' && student.charAt(i+2)=='t' && student.charAt(i+3)=='%')
//                     break;
//                 
//                 if( forend && (int)student.charAt(j)>47 && (int)student.charAt(j)<58 )
//                 {
//                    endfactor=j;
//                    forend=false;
//                 }
//                 i--;
//                 j--; 
//             } 
//           //  System.out.println(student.substring(j,endfactor+1));
//             startfactor=j-subjectline.start();
//             endfactor=endfactor-subjectline.start();
//          //  System.out.println("parameter ="+startfactor+"  " +endfactor);
//             
//                         
//        }
//    public void  primarycontroller(File file,int year,boolean sem2) throws IOException 
//    { 
//           pdffile =file;
//          try{   
//                 XSSFWorkbook workbook = new XSSFWorkbook();  
//	         XSSFSheet sheet = workbook.createSheet("student Marksheet"); 
//                // XSSFSheet sheet1 = workbook.createSheet("Congratulations"); 
//	         PDDocument document = PDDocument.load(pdffile); 
//	         if (!document.isEncrypted()) 
//	         {
//		          PDFTextStripper stripper = new PDFTextStripper();  
//		          String text = stripper.getText(document);
//		        
//		          String student[]=text.split("TOTAL CREDITS E");
//		          int rownum = 0;  
//		          int cellnum=0; 
//                          System.out.println(student[1]);
//		          init(year,text,cellnum,rownum,sheet,workbook,student[1]); 
//		          rownum=7;
//		          int no=1; 
//                          gocheck(student[0],year);
//		          for(int i=0;i<student.length-1;i++)
//		          {			        	 
//			        rownum=getSeatNoAndName(year,student[i],workbook,sheet,rownum,no); 
//                                //   System.out.println("after name and no");
//                                int  earnedcredit =putearnedcredit(student[i+1],workbook,sheet,rownum,sem2end); 
//                                //System.out.println("after earned credit");
//                                excelsgpa(workbook,sheet,rownum,sem2end+1,year);
//                               // System.out.println("after excel sgpa");
//                                if(year==4)
//                                    sgpalastyear(student[i],workbook,sheet,rownum,sem2end+2);
//                                else
//                                pdfsgpa(student[i],workbook,sheet,rownum,sem2end+2); 
//                                finalresult(workbook,sheet,rownum,sem2end+3,earnedcredit,year);
//                                if(year==4)
//                                {
//                                    sgpacgpa(student[i+1],workbook,sheet, rownum,sem2end+4);
//                                }
//			        no++;  
//		          }
//                          
//		        for (int i=1; i<3; i++)
//                        {
//		           sheet.autoSizeColumn(i);
//		        }
//		        for (int i=sem2end+1; i<sem2end+5; i++)
//                        {
//		           sheet.autoSizeColumn(i);
//  	                   sheet.addMergedRegion(new CellRangeAddress(5,6,i,i)); 
//		        }
//                        FileOutputStream out = new FileOutputStream("result_"+year+".xlsx"); 
//                        workbook.write(out); 
//                        out.close(); 
//                        System.out.println("result.xlsx written successfully on disk.");    
//	         } 
//	         workbook.close();
//	         document.close();
//          }
//          catch(Exception e)
//         {
//            System.out.println("exception");
//            System.out.println(e.toString());  
//         }  
//    } 
//    public static void  sgpacgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
//    {
//        fesgpa(text,workbook,sheet , rownum,cellno);
//        sesgpa(text,workbook,sheet , rownum,cellno+2);
//        tesgpa(text,workbook,sheet , rownum,cellno+3);
//        String statement =cgpa(text,workbook,sheet , rownum,cellno+4);
//        finalstatement(text,workbook,sheet , rownum,cellno+5,statement);
//        
//    }
//    public static void finalstatement(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno,String degree)
//    {
//        CellStyle style = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//          
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno); 
//        try{  
//             cell.setCellStyle(style);
//             cell.setCellValue(degree);  
//        }
//        catch(Exception e)
//        {
//             System.out.println("exception");
//            System.out.println(e.toString()); 
//        }
//    }
//    public static void fesgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
//    {
//        CellStyle style = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//          
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno);
//        try{
//         Pattern sgpap=Pattern.compile("FE SGPA :");
//          Matcher m=sgpap.matcher(text); 
//          if(m.find())
//          {
//              String sgpa=text.substring(m.end(),m.end()+6); 
//              sgpa= sgpa.trim();
//             
//             row = sheet.getRow(rownum-1);   
//             cell = row.createCell(++cellno); 
//             cell.setCellStyle(style);
//            cell.setCellValue(sgpa); 
//            
//          }
//        }
//        catch(Exception e)
//        {
//             System.out.println("exception");
//            System.out.println(e.toString());
//            return;
//        }
//    }
//    public static void sesgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
//    {
//        CellStyle style = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        
//        try{
//         Pattern sgpap=Pattern.compile("SE SGPA :");
//          Matcher m=sgpap.matcher(text); 
//          if(m.find())
//          {
//              String sgpa=text.substring(m.end(),m.end()+5); 
//              sgpa= sgpa.trim();
//             
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno); 
//             cell.setCellStyle(style);
//            cell.setCellValue(sgpa); 
//            
//          }
//        }
//        catch(Exception e)
//        {
//             System.out.println("exception");
//            System.out.println(e.toString());
//            return;
//        }
//    }
//    public static void tesgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
//    {
//        CellStyle style = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        
//        try{
//         Pattern sgpap=Pattern.compile("TE SGPA :");
//          Matcher m=sgpap.matcher(text); 
//          if(m.find())
//          {
//              String sgpa=text.substring(m.end(),m.end()+5); 
//              sgpa= sgpa.trim();
//             
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno); 
//             cell.setCellStyle(style);
//            cell.setCellValue(sgpa); 
//            
//          }
//        }
//        catch(Exception e)
//        {
//             System.out.println("exception");
//            System.out.println(e.toString()); 
//        }
//    }
//    public static String cgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
//    {
//        CellStyle style = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        
//        try{
//         Pattern sgpap=Pattern.compile("CGPA :");
//          Matcher m=sgpap.matcher(text); 
//          if(m.find())
//          {
//              String cgpa=text.substring(m.end(),m.end()+5); 
//              cgpa= cgpa.trim();
//             
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno); 
//            cell.setCellStyle(style);
//            cell.setCellValue(cgpa);
//            String degree=text.substring(m.end()+5);
//              int i=0;
//              while(degree.charAt(i)!='\n')
//              {
//                  i++;
//              }
//              
//             degree=degree.substring(0, i);
//             return degree;
//            
//          }
//        }
//        catch(Exception e)
//        {
//             System.out.println("exception");
//            System.out.println(e.toString());
//            return "RESULT RESERVED FOR BACKLOGS";
//        }
//        return "RESULT RESERVED FOR BACKLOGS";
//    }
//
//    public static void finalresult(XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno,int credits,int y)
//    {
//        //RRB part is remaaning //fc part
//           CellStyle style = workbook.createCellStyle();
//          style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        int totalcredits=-1;
//        switch(y)
//        {
//            case 2:
//                totalcredits=50;
//            case 3:
//                totalcredits=46;
//            case 4:
//                totalcredits=44;
//                
//        } 
//        if(credits==-1)
//            return;
//        if(credits<totalcredits)
//        {
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno);   
//             cell.setCellStyle(style);
//            cell.setCellValue("ATKT"); 
//        }
//        else
//        {
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno); 
//             cell.setCellStyle(style);
//            cell.setCellValue("FD");  
//        }
//    }
//    public static void excelsgpa(XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno, int y)
//    {
//        try
//        {
//        CellStyle style = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        
//        int sem1cellno=sem1end,sem2cellno=sem2end-1;
//        Row row=sheet.getRow(rownum-1); 
//        Cell cell = row.createCell(++cellno);
//         int totalcredits=1;
//        int sem1=(int)row.getCell(sem1cellno).getNumericCellValue();
//        int semm2=(int)row.getCell(sem2cellno).getNumericCellValue();
//      //  System.out.print("sem1: "+sem1+" sem2: "+semm2);
//        switch(y)
//        {
//            case 2:
//                totalcredits=50;
//                break;
//            case 3:
//                totalcredits= 46;
//                break;
//            case 4:
//                totalcredits=44;
//                break;
//                
//        }  
//        double sgpa=(double)(sem1+semm2)/totalcredits;
//        sgpa = Math.round(sgpa * 100.0) / 100.0;
//        System.out.println(" sgpa: "+ sgpa);
//        cell.setCellStyle(style);
//        cell.setCellValue(sgpa); 
//        }
//        catch(Exception e)
//        {
//            System.out.println("sgpa e");
//            return;
//        }
//        
//    }
//    public static void sgpalastyear(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
//    {
//        CellStyle style = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        
//        
//         Pattern sgpap=Pattern.compile("FOURTH YEAR SGPA : ");
//          Matcher m=sgpap.matcher(text); 
//          if(m.find())
//          {
//              String sgpa=text.substring(m.end(),text.length());
//              sgpa=sgpa.replace(',', ' ');
//              sgpa= sgpa.trim();
//             
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno); 
//             cell.setCellStyle(style);
//            cell.setCellValue(sgpa); 
//            
//          }
//    }
//    public static void pdfsgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
//    {
//         CellStyle style = workbook.createCellStyle();
//          style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        
//        
//         Pattern sgpap=Pattern.compile("SGPA : ");
//          Matcher m=sgpap.matcher(text); 
//          if(m.find())
//          {
//              String sgpa=text.substring(m.end(),text.length());
//              sgpa=sgpa.replace(',', ' ');
//              sgpa= sgpa.trim();
//             
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(++cellno); 
//             cell.setCellStyle(style);
//            cell.setCellValue(sgpa); 
//            
//          }
//    }
//     public static int putearnedcredit(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
//     {
//          CellStyle style = workbook.createCellStyle();
//          style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//        
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        
//        
//         //extracted form pdf
//          Pattern earn=Pattern.compile("ARNED :" );
//          Matcher m=earn.matcher(text);  
//          if(m.find())
//          {
//          String totalcreditsearned =text.substring(m.end(),m.end()+3);
//         totalcreditsearned= totalcreditsearned.replaceAll("[^0-9]", "");
//          
//            Row row = sheet.getRow(rownum-1);   
//            Cell cell = row.createCell(cellno);  
//            
//            cellno++;
//            cell=row.createCell(cellno);
// cell.setCellStyle(style);
//            cell.setCellValue(totalcreditsearned); 
//            System.out.println(totalcreditsearned);
//
//             return Integer.parseInt(totalcreditsearned);
//          }
//          return -1;
//     }
//   public static void init(int year ,String text,int cellnum,int rownum,XSSFSheet sheet,XSSFWorkbook workbook,String student) {
//        
//        int k=0,n=0;
//        getSubjectCode(year,student,sheet,workbook);
//             switch (year) {
//             case 2:
//                 k=6;n=12;
//                 break;
//             case 3:
//                 k=8; n=16;
//                 break;
//             case 4:
//                 k=8;n=15;
//                 break;
//         }
//         
//            sem1end=k;sem2end=n; 
//             System.out.println("parameter sem="+sem1end+"  " +sem2end);
//             for(int i=0;i<n;i++){
//                  System.out.println("subn =i="+i+" "+subn[i]);
//                 if(subn[i]==1){
//                     if(i<k){
//                          sem1end++;
//                     }
//                     sem2end++;
//                 }
//             }
//             System.out.println("parameter sem="+sem1end+"  " +sem2end);
//                sem1end=(sem1end*5)+3;
//                sem2end=(sem2end*5)+6;
//             
//                 System.out.println("parameter sem="+sem1end+"  " +sem2end);
//                
//    	
//    	CellStyle style = workbook.createCellStyle(); 
//    	CellStyle style1 = workbook.createCellStyle();
//        style1.setAlignment(CellStyle.ALIGN_CENTER);
//        style1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        Font font = workbook.createFont();  
//        Font font1 = workbook.createFont();  
//        font1.setFontHeightInPoints((short)20);  
//        font1.setFontName("Courier New");  
//        
//        font.setBold(true);  
//        font1.setBold(true); 
//        // Applying font to the style  
//        style1.setFont(font1); 
//       // style.setFont(font); 
//        
//        style1.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
//        
//        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
//    	Row row0 = sheet.createRow(1); 
//    	row0.setHeight((short)500);
//    	cellnum=2;
//    	Cell cell1 = row0.createCell(cellnum++); 
//        
//        cell1.setCellValue("\t\t   RESULT"); 
//        
//        style1.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
//        style1.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style1.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
//        style1.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style1.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
//        style1.setTopBorderColor(IndexedColors.BLACK.getIndex());
//
//        cell1.setCellStyle(style1);
//        for(int i=0;i<sem2end+3;i++) {
//        cell1 = row0.createCell(cellnum++); 
//        cell1.setCellStyle(style1);
//        }
////        ---------------------------------
//        style1.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
//        style1.setRightBorderColor(IndexedColors.BLACK.getIndex());  
//        
//        cell1.setCellStyle(style1);
//        CellStyle style2 = workbook.createCellStyle();
//        style2.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
//        style2.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style2.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
//        style2.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style2.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
//        style2.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style2.setFont(font); 
//        style2.setAlignment(CellStyle.ALIGN_CENTER);
//        style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        style2.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());  
//        
//        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        sheet.addMergedRegion(new CellRangeAddress(1,1,2,sem2end+5));  
//        sheet.addMergedRegion(new CellRangeAddress(3,3,3,sem1end));  
//        sheet.addMergedRegion(new CellRangeAddress(3,3,sem1end+2,sem2end));  
//        sheet.addMergedRegion(new CellRangeAddress(5,6,0,0)); 
//	sheet.addMergedRegion(new CellRangeAddress(5,6,1,1)); 
//	
//
//		rownum=3;
//        row0 = sheet.createRow(rownum++); 
//        cellnum=0;
//        row0.setHeight((short)500);
//        cellnum=3;
//        cell1 = row0.createCell(cellnum++); 
//        
//        cell1.setCellValue("\t\tSEM 1"); 
//        style1.setFont(font1); 
//        cell1.setCellStyle(style1);
//        for(int i=0;i<sem1end-3;i++) {
//            cell1 = row0.createCell(cellnum++); 
//            cell1.setCellStyle(style1);
//            }
//         
//        style1.setFont(font1); 
//        cell1.setCellStyle(style1);
//        cell1 = row0.createCell(cellnum++); 
//        cell1 = row0.createCell(cellnum++);
//        cell1.setCellValue("\t\tSEM 2");
//        cell1.setCellStyle(style1);
//        for(int i=0;i<sem2end-sem1end-2;i++) {
//            cell1 = row0.createCell(cellnum++); 
//            cell1.setCellStyle(style1);
//            } 
//        rownum=5;
//    	row0 = sheet.createRow(rownum++); 
//    	 row0.setHeight((short)400);
//    	cellnum=0;
//        Cell cell0 = row0.createCell(cellnum++); 
//        cell0.setCellValue("  Sr No.  "); 
//        cell0.setCellStyle(style2);
//        cell0 = row0.createCell(cellnum++); 
//        cell0.setCellValue("Seat no"); 
//        cell0.setCellStyle(style2);
//        cell0=row0.createCell(cellnum++);
//        cell0.setCellValue("Subject Code");
//        cell0.setCellStyle(style2);
//        
//        // Setting Foreground Color  
//        style = workbook.createCellStyle(); 
//        
//        style.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setFont(font); 
//       // style.setFillForegroundColor(IndexedColors.BLUE.getIndex());  
//       // style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//
//        style.setAlignment(CellStyle.ALIGN_CENTER);
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//   
//        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());  
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//    //    style.setVerticalAlignment(VerticalAlignment.CENTER);
//        int cm=3;
//         
////                System.out.println("sun= 11 "+sub[11]);
////                System.out.println("sun=12 "+sub[12]); 
//       // System.out.println("sub l "+sub.length); 
//        for (int i=1; i<=n; i++){
//            
//             if(sub[i-1]==null)
//                break;
////            System.out.println("sun="+(i)+" "+subn[i-1]); 
////        	System.out.println("sun="+(i)+" "+sub[i-1]);
//                
//                
//        	cell0=row0.createCell(cellnum++);
//        	cell0.setCellValue(sub[i-1]);
//       	    cell0.setCellStyle(style);
//        	 if(subn[i-1]==0) {
//                     if(i==k+1) {
//        			 cm=cm+2;
//        		 }
//        		 sheet.addMergedRegion(new CellRangeAddress(5,5,cm,cm+4)); 
//        		 cm=cm+5;
//        		 for (int j=0; j<4; j++){
//             		
//            		 cell0=row0.createCell(cellnum++);
//            		 cell0.setCellStyle(style);
//            	 }
//        		 
//        	 }
//        	 else {
//        		 if(i==k+1) {
//        			 cm=cm+2;
//        		 }
//        		 sheet.addMergedRegion(new CellRangeAddress(5,5,cm,cm+9)); 
//        		 cm=cm+10;
//        		 for (int j=0; j<9; j++){
//             		
//            		 cell0=row0.createCell(cellnum++);
//            		 cell0.setCellStyle(style);
//            	 }
//        		
//        	 }
//        	
//        	 cell0.setCellStyle(style);
//        	 if(i==k||i==n) {
//        		 cell0=row0.createCell(cellnum++);
//
//        	        style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
//        	        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        	    	cell0.setCellValue("Sem "+(i/6));
//        	        cell0.setCellStyle(style); 
//        	        cell0=row0.createCell(cellnum++);
//        	    	cell0.setCellValue("        ");
//        	   
//        	        
//        	 }
//        	
//     	}
//        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
//        
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        if(year==4){
//             cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("Earned Credits");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("SGPA");
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("SGPA (PDF)");
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("Result");
//        cell0.setCellStyle(style);
//               cell0=row0.createCell(cellnum++);
//               cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("FE SGPA");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("SE SGPA");
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("TE SGPA");
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("CGPA");
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("Final Result");
//        cell0.setCellStyle(style);
//        }
//        else
//        {     cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("Earned Credits");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("SGPA");
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("SGPA (PDF)");
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    	cell0.setCellValue("Result");
//        cell0.setCellStyle(style);
//        }
// //       ===============================================================================
//        
//        row0 = sheet.createRow(rownum++); 
//        row0.setHeight((short)400);
//        cell0=row0.createCell(2);
//        cell0.setCellValue("Name");
//        cell0.setCellStyle(style); 
//        cellnum=3;
//        style = workbook.createCellStyle(); 
//        style.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//          style.setAlignment(CellStyle.ALIGN_CENTER); 
//          style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//
//        style.setFillForegroundColor(IndexedColors.TAN.getIndex());  
//        
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setFont(font); 
//        cell0=row0.createCell(0);
//        
//   	    cell0.setCellStyle(style);
//   	 cell0=row0.createCell(1);
//	    cell0.setCellStyle(style);
//   		for (int j=0; j<n; j++){
//   		 cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("  TOT  ");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("  Crd  ");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("Grd");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("Grdpts");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("Crdpts");
//        cell0.setCellStyle(style);
//        if(subn[j]==1){
//        	cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("TW");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("Crd");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("Grd");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("Grdpts");
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//   		cell0.setCellValue("Crdpts");
//        cell0.setCellStyle(style);
//        }
//        
//        if(j==k-1||j==n-1) {
//        	CellStyle style3 = workbook.createCellStyle(); 
//            style3.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
//            style3.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//            style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style3.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
//            style3.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//            style3.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
//            style3.setTopBorderColor(IndexedColors.BLACK.getIndex());
//            style3.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
//            style3.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//              style3.setAlignment(CellStyle.ALIGN_CENTER); 
//              style3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        	style3.setFillForegroundColor(IndexedColors.BLUE.getIndex());  
//            style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        	cell0=row0.createCell(cellnum++);
//       		cell0.setCellValue("  TOTCrdpts  ");
//            cell0.setCellStyle(style3);
//            cell0=row0.createCell(cellnum++);
//       		cell0.setCellValue("      ");
//           
//        
//        }
//        style.setFillForegroundColor(IndexedColors.TAN.getIndex());  
//        
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//   	 }
////style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
//        
////        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        cell0=row0.createCell(cellnum++);
//    	
//        cell0.setCellStyle(style); 
//        cell0=row0.createCell(cellnum++);
//    	
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    
//        cell0.setCellStyle(style);
//        cell0=row0.createCell(cellnum++);
//    	
//        cell0.setCellStyle(style);
//   	 
//    }
//     public static void getSubjectCode(int y,String text, XSSFSheet sheet,XSSFWorkbook workbook)
//    {
//         String year="";
//          switch (y) {
//             case 2:
//                 year= "20";
//                 break;
//             case 3:
//                 year = "30";
//                 break;
//             case 4:
//                 year = "40";
//                 break;
//         }
//    	  Pattern subject=Pattern.compile(year+"[0-9]{4}" );
//          Matcher m=subject.matcher(text); 
//          Pattern sem2p=Pattern.compile("SEM.:2" );
//          Matcher m2=subject.matcher(text);  
//          if(m2.find())
//             sem2=true;
//          else
//             sem2=false; 
//          String prev="";
//          int i=0;
//          if(m.find())
//              prev=text.substring(m.start(),m.end());
//          sub[i]=prev;
//          while(m.find())
//          {  
//            String temp=text.substring(m.start(),m.end());
//            if(temp.equals(prev))
//                subn[i]=1;
//            else
//            {
//                i++;
//                sub[i]=temp;
//                prev=temp;
//            }  
//          } 
//    } 
//    public static void putSubjectData(XSSFWorkbook workbook,String[] studentdata, XSSFSheet sheet, Row row,int cellno,int y,int sem1,int sem2)
//    { 
//    	 CellStyle style = workbook.createCellStyle();
//          //style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());   
//        //style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//        style.setAlignment(CellStyle.ALIGN_CENTER); 
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//         int sem1position=0, sem2position=0;
//         switch(y)
//         {
//             case 2:
//             {
//                 sem1position=58;
//                 sem2position=110;
//                 break;
//             }
//             case 3:
//             {
//                 sem1position=53;
//                 sem2position=105;
//                 break;
//             }
//             case 4:
//             {
//                 sem1position=53;
//                 sem2position=100;
//                 break;
//             }
//                 
//         }
//    	for(int i=0; studentdata[i]!=null && i<studentdata.length;i++)
//    	{
//            if(cellno==sem1position) 
//            { 
//                CellStyle style3 = workbook.createCellStyle(); 
//                style3.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
//                style3.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//                style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
//                style3.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
//                style3.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//                style3.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
//                style3.setTopBorderColor(IndexedColors.BLACK.getIndex());
//                style3.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
//                style3.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//                style3.setAlignment(CellStyle.ALIGN_CENTER); 
//                style3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//                style3.setFillForegroundColor(IndexedColors.BLUE.getIndex());  
//                style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
//
//                Cell cell=row.createCell(cellno++);
//                cell.setCellStyle(style3);                        
//                cell.setCellValue(sem1); 
//                cellno++;  
//            }
//            
//            Cell cell = row.createCell(cellno++); 
//            cell.setCellValue(studentdata[i]);
//            cell.setCellStyle(style); 
//            
//            if(cellno==sem2position)
//            {
//                CellStyle style3 = workbook.createCellStyle(); 
//                style3.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
//                style3.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//                style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
//                style3.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
//                style3.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//                style3.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
//                style3.setTopBorderColor(IndexedColors.BLACK.getIndex());
//                style3.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
//                style3.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//                style3.setAlignment(CellStyle.ALIGN_CENTER); 
//                style3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//                style3.setFillForegroundColor(IndexedColors.BLUE.getIndex());  
//                style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
//
//                cell=row.createCell(cellno++);
//                cell.setCellStyle(style3);
//                cell.setCellValue(sem2); 
//                cellno++;
//            }
//    	}
//    } 
//    public static String[] getSubjectData(String text) throws InterruptedException
//    {  
//    	
//        int size=0;
//        for(int i=0;i<sub.length;i++)
//        {
//            if(subn[i]==1)
//                size+=10;
//            else if(sub[i]!=null)
//                size+=5;
//        }
//    	String studentdata[]=new String[size+11]; 
//    	int j=0;
//    	for(int k=0;k<sub.length;k++)
//        {
//            if(sub[k]==null)
//                break;
//    	    Pattern subjectcode=Pattern.compile(sub[k]);            
//            Matcher sub=subjectcode.matcher(text);      
//            int end;
//            while(sub.find())
//            { 
//                int  start=sub.start()+startfactor;
//                if((int)text.charAt(sub.end())>64 && ((int)text.charAt(sub.end())<70) )
//                {
//                     
//                   end=sub.start()+endfactor+2;
//                }
//                else
//               end=sub.start()+endfactor+1;
//                String marks=text.substring(start,end);
//                String arr[]=marks.split(" ");  
//                 for(int i=0;i<arr.length;i++)
//                 { 
//                    if(arr[i].length()>=1 && arr[i]!=null)
//                    {
//                        studentdata[j]=arr[i];
//                        j++;
//                    }
//                 }  
//            } 
//    	}  
//	return studentdata;
//    } 
//    public static int  getSeatNoAndName(int year ,String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int no) throws InterruptedException
//    {
//    	  
//    	  Pattern seatNo;
//         switch (year) {
//             case 2:
//                 seatNo = Pattern.compile("S"+"[0-9]{9}" );
//                 break;
//             case 3:
//                 seatNo = Pattern.compile("T"+"[0-9]{9}" );
//                 break;
//             default:
//                 seatNo = Pattern.compile("B"+"[0-9]{9}" );
//                 break;
//         }
//
//          Matcher m=seatNo.matcher(text); 
//          int cellnum=0;
////          Row row0 = sheet.createRow(rownum++); 
////          Cell cell0 = row0.createCell(cellnum++); 
////          cell0.setCellValue("Seat no"); 
////          cell0=row0.createCell(cellnum++);
////          cell0.setCellValue("Name"); 
//          while(m.find())
//          {  
//             // cellnum=0;
//        	CellStyle style = workbook.createCellStyle();  
//        	style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
//        	style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//                style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//                style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//                style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//                style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//                style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//                style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//                style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//                style.setAlignment(CellStyle.ALIGN_CENTER); 
//                style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); 
//                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//                style.setAlignment(CellStyle.ALIGN_CENTER);  
//                style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); 
//              Row row = sheet.createRow(rownum++);     
//              Cell cell = row.createCell(cellnum++); 
//              cell.setCellStyle(style);
//              cell.setCellValue(no); 
//              cell = row.createCell(cellnum++); 
//              cell.setCellValue(text.substring(m.start(), m.end())); 
//              
//              String name=getName(text.substring(m.end()+2,m.end()+50));
//              
//             System.out.print(no+") "+ name+"  ");
//              cell.setCellStyle(style);
//              cell=row.createCell(cellnum++);
//              cell.setCellValue(name); 
//               style = workbook.createCellStyle();
//               style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
////             
//             style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//             style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//             style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
//             style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//             style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
//             style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
//             style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
//             style.setTopBorderColor(IndexedColors.BLACK.getIndex());
//             style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
//             style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
//             style.setAlignment(CellStyle.ALIGN_CENTER); 
//             style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//             cell.setCellStyle(style); 
//            
//              String studentdata[]=getSubjectData(text);
//               
//              int sem_marks[]=new int[2];
//              sem_marks=calculate_semmarks(studentdata,year);
//             
//            //  System.out.println(sem1+" "+ sem2);
//            //  Thread.sleep(1);
//              putSubjectData(workbook,studentdata,sheet,row,cellnum,year,sem_marks[0],sem_marks[1]);
//                 
//          } 
//          return rownum;
//    }
//    public static int[] calculate_semmarks(String studentdata[],int y )
//    {
//         int count=0;
//         int arr[]=new int[2];
//         int semdivider=0;
//         switch(y){
//             case 2:
//                 semdivider=11;
//                 break;
//             case 3:
//                 semdivider=10;
//                 break;
//             case 4:
//                 semdivider=10;
//                 break;
//         }  
//              for(int i=1;i<studentdata.length;i++)
//              { 
//                  if(studentdata[i-1]!=null && i%5==0)
//                  {
//                      if(count<semdivider)
//                      {
//                          arr[0]+=Integer.parseInt(studentdata[i-1]);
//                          count++;
//                      }
//                      else
//                      {
//                            arr[1]+=Integer.parseInt(studentdata[i-1]);
//                            count++;
//                      }
//                  }
//              }
//              return arr;
//    }
//    public static String getName(String text)
//    {
//        int start=0;
//    	if(text.charAt(start)==' ')
//    	start++;
//        
//    	String name="";  
//    	while(true)
//    	{
//    		if(text.charAt(start)==' ' && text.charAt(start+1)==' ' || start==47)
//    			break;
//    		name+=text.charAt(start);
//    		start++;    				
//    	}    
//    	return name;
//    }
//    
//}
//package com.SmartResult;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.regex.*;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;  
import org.apache.poi.ss.usermodel.IndexedColors;  
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;  
import org.apache.poi.ss.usermodel.*;  
import org.apache.poi.xssf.usermodel.XSSFCellStyle;  
class Converter{
    File pdffile;
    public static int subn[]=new int[22];
    public static String sub[]=new String[22];
     public static int startfactor,endfactor,sem1end,sem2end,numof_students=0;
     public static boolean sem2=false;
     public static void gocheck(String student,int y)
    {
         Pattern total=Pattern.compile("Tot%");
         Matcher t=total.matcher(student);
         String year="";
         switch (y) {
         case 2:
             year= "20";
             break;
         case 3:
             year = "30";
             break;
         case 4:
             year = "40";
             break;
        }
         Pattern subject=Pattern.compile(year+"4181") ;
         Matcher subjectline=subject.matcher(student);
         t.find();
         subjectline.find();
         int i=t.end();
         while(student.charAt(i+1)!='\n')// for total
         {
             i++;
         }
         int j=subjectline.end();
         while(student.charAt(j+1)!='\n')  //for subjectline
         {
             j++;
         }
        // System.out.println(student.substring(subjectline.start(), j));
         int end=j;
         boolean forend=true;
         while(true)
         {
             if(student.charAt(i)=='T' && student.charAt(i+1)=='o' && student.charAt(i+2)=='t' && student.charAt(i+3)=='%')
                 break;

             if( forend && (int)student.charAt(j)>47 && (int)student.charAt(j)<58 )
             {
                endfactor=j;
                forend=false;
             }
             i--;
             j--; 
         } 
         System.out.println(student.substring(j,endfactor+1));
         startfactor=j-subjectline.start();
         endfactor=endfactor-subjectline.start();
        System.out.println("parameter ="+startfactor+"  " +endfactor);


    }
    public void  primarycontroller(File file,int year,int sem) throws IOException 
    { 
                pdffile =file;
        try{   
                 XSSFWorkbook workbook = new XSSFWorkbook();  
	         XSSFSheet sheet = workbook.createSheet("student Marksheet"); 
                 XSSFSheet sheet1 = workbook.createSheet("Congratulations"); 
	         PDDocument document = PDDocument.load(pdffile); 
//                 System.out.println("sem="+sem);
                 if(sem == 0){
//                     System.out.println("cellno= false");
                     sem2=false;
                 }
                 else{
                     sem2=true;
                 }
	         if (!document.isEncrypted()) 
	         {
                    PDFTextStripper stripper = new PDFTextStripper();  
                    String text = stripper.getText(document);

                    String student[]=text.split("TOTAL CREDITS E");
                    int rownum = 0;  
                    int cellnum=0; 
                    System.out.println(student[1]);
                    init(year,sem,text,cellnum,rownum,sheet,workbook,student[1]); 
                    rownum=7;
                    int no=1;
                 //   System.out.println("init succesfully executed"+" sem2="+sem2);
                    gocheck(student[0],year);

                    for(int i=0;i<student.length-1;i++)
                    {			        	 
                          rownum=getSeatNoAndName(year,student[i],workbook,sheet,rownum,no); 
                          //   System.out.println("after name and no");
                          int  earnedcredit =putearnedcredit(student[i+1],workbook,sheet,rownum,sem2end); 
                          //System.out.println("after earned credit");
                          if(sem2){
                              excelsgpa(workbook,sheet,rownum,sem2end+1,year);
                          }
                         // System.out.println("after excel sgpa");
                          if(year==4)
                              sgpalastyear(student[i],workbook,sheet,rownum,sem2end+2);
                          else
                          pdfsgpa(student[i],workbook,sheet,rownum,sem2end+2); 
                          finalresult(workbook,sheet,rownum,sem2end+3,earnedcredit,year);
                          if(year==4)
                          {
                              sgpacgpa(student[i+1],workbook,sheet, rownum,sem2end+4);
                          }
                          no++;  
                    }
               // this Writes the workbook  
                  for (int i=1; i<3; i++)
                  {
                     sheet.autoSizeColumn(i);
                  }
                  for (int i=sem2end+1; i<sem2end+5; i++)
                  {
                     sheet.autoSizeColumn(i);
                     sheet.addMergedRegion(new CellRangeAddress(5,6,i,i)); 
                  }
//                        System.out.println("subtop");

                  getSubTop(workbook, sheet,year,numof_students);
                  getSubCount(workbook, sheet,year,numof_students);
                  if(sem2){
                  putCongdata(workbook, sheet,sheet1,year,numof_students);
                  getOverAlldata(workbook, sheet,sheet1,year,numof_students);
                  putOverAlldata(workbook, sheet,sheet1,year,numof_students);
                  getTopperdata(workbook, sheet,sheet1,year,numof_students);

                  for (int i=3; i<7; i++)
                  {
                     sheet1.autoSizeColumn(i);
                  }
                  for (int i=9; i<11; i++)
                  {
                     sheet1.autoSizeColumn(i);
                  }
                  for (int i=13; i<16; i++)
                  {
                     sheet1.autoSizeColumn(i);
                  }
                   }
                  FileOutputStream out = new FileOutputStream("result_year_"+year+"_sem_"+(sem+1)+".xlsx"); 
                  workbook.write(out); 
                  out.close(); 
                  System.out.println("result.xlsx written successfully on disk.");    
	         } 
	         workbook.close();
	         document.close();
          }
          catch(Exception e)
         {
            System.out.println("exception");
            System.out.println(e.toString());  
         }  
    } 
    public static void fesgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
    {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
          
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno);
        try{
         Pattern sgpap=Pattern.compile("FE SGPA :");
          Matcher m=sgpap.matcher(text); 
          if(m.find())
          {
              String sgpa=text.substring(m.end(),m.end()+6); 
              sgpa= sgpa.trim();
             
             row = sheet.getRow(rownum-1);   
             cell = row.createCell(++cellno); 
             cell.setCellStyle(style);
            cell.setCellValue(sgpa); 
            
          }
        }
        catch(Exception e)
        {
             System.out.println("exception");
            System.out.println(e.toString());
            return;
        }
    }
    public static void sesgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
    {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        
        try{
         Pattern sgpap=Pattern.compile("SE SGPA :");
          Matcher m=sgpap.matcher(text); 
          if(m.find())
          {
              String sgpa=text.substring(m.end(),m.end()+5); 
              sgpa= sgpa.trim();
             
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno); 
             cell.setCellStyle(style);
            cell.setCellValue(sgpa); 
            
          }
        }
        catch(Exception e)
        {
             System.out.println("exception");
            System.out.println(e.toString());
            return;
        }
    }
    public static void tesgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
    {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        
        try{
         Pattern sgpap=Pattern.compile("TE SGPA :");
          Matcher m=sgpap.matcher(text); 
          if(m.find())
          {
              String sgpa=text.substring(m.end(),m.end()+5); 
              sgpa= sgpa.trim();
             
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno); 
             cell.setCellStyle(style);
            cell.setCellValue(sgpa); 
            
          }
        }
        catch(Exception e)
        {
             System.out.println("exception");
            System.out.println(e.toString());
            return;
        }
    }
    public static String cgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
    {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        
        try{
         Pattern sgpap=Pattern.compile("CGPA :");
          Matcher m=sgpap.matcher(text); 
          if(m.find())
          {
              String cgpa=text.substring(m.end(),m.end()+5); 
              cgpa= cgpa.trim();
             
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno); 
            cell.setCellStyle(style);
            cell.setCellValue(cgpa);
            String degree=text.substring(m.end()+5);
              int i=0;
              while(degree.charAt(i)!='\n')
              {
                  i++;
              }
              
             degree=degree.substring(0, i);
             return degree;
            
          }
        }
        catch(Exception e)
        {
             System.out.println("exception");
            System.out.println(e.toString());
            return "RESULT RESERVED FOR BACKLOGS";
        }
        return "RESULT RESERVED FOR BACKLOGS";
    }
public static void finalresult(XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno,int credits,int y)
    {
        //RRB part is remaaning
           CellStyle style = workbook.createCellStyle();
          style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        int totalcredits=-1;
        switch(y)
        {
            case 2:
                totalcredits=50;
            case 3:
                totalcredits=46;
            case 4:
                totalcredits=44;
                
        } 
        if(credits==-1)
            return;
        if(credits<totalcredits)
        {
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno);   
             cell.setCellStyle(style);
            cell.setCellValue("ATKT"); 
        }
        else
        {
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno); 
             cell.setCellStyle(style);
            cell.setCellValue("FD");  
        }
    }
   public static void excelsgpa(XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno, int y)
    {
        try
        {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        
        int sem1cellno=sem1end,sem2cellno=sem2end-1;
        Row row=sheet.getRow(rownum-1); 
        Cell cell = row.createCell(++cellno);
         int totalcredits=1;
        int sem1=(int)row.getCell(sem1cellno).getNumericCellValue();
        int semm2=(int)row.getCell(sem2cellno).getNumericCellValue();
      //  System.out.print("sem1: "+sem1+" sem2: "+semm2);
        switch(y)
        {
            case 2:
                totalcredits=50;
                break;
            case 3:
                totalcredits= 46;
                break;
            case 4:
                totalcredits=44;
                break;
                
        } 
        
        double sgpa=(double)(sem1+semm2)/totalcredits;
        sgpa = Math.round(sgpa * 100.0) / 100.0;
      //  System.out.println(" sgpa: "+ sgpa);
        cell.setCellStyle(style);
        cell.setCellValue(sgpa); 
        }
        catch(Exception e)
        {
            System.out.println("sgpa e");
            return;
        }
        
    }
   public static void sgpalastyear(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
    {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        
        
         Pattern sgpap=Pattern.compile("FOURTH YEAR SGPA : ");
          Matcher m=sgpap.matcher(text); 
          if(m.find())
          {
              String sgpa=text.substring(m.end(),text.length());
              sgpa=sgpa.replace(',', ' ');
              sgpa= sgpa.trim();
             
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno); 
             cell.setCellStyle(style);
            cell.setCellValue(sgpa); 
            
          }
    }
    public static void pdfsgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
    {
         CellStyle style = workbook.createCellStyle();
          style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        
        
         Pattern sgpap=Pattern.compile("SGPA : ");
          Matcher m=sgpap.matcher(text); 
          if(m.find())
          {
              String sgpa=text.substring(m.end(),text.length());
              sgpa=sgpa.replace(',', ' ');
              sgpa= sgpa.trim();
             
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno); 
             cell.setCellStyle(style);
           
                 cell.setCellValue(sgpa); 
            
            
          }
    }
    public static void finalstatement(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno,String degree)
    {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
          
            Row row = sheet.getRow(rownum-1);   
            Cell cell = row.createCell(++cellno); 
        try{  
             cell.setCellStyle(style);
             cell.setCellValue(degree);  
        }
        catch(Exception e)
        {
             System.out.println("exception");
            System.out.println(e.toString()); 
        }
    }
    
    public static void  sgpacgpa(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
    {
        fesgpa(text,workbook,sheet , rownum,cellno);
        sesgpa(text,workbook,sheet , rownum,cellno+2);
        tesgpa(text,workbook,sheet , rownum,cellno+3);
        String statement =cgpa(text,workbook,sheet , rownum,cellno+4);
        finalstatement(text,workbook,sheet , rownum,cellno+5,statement);
        
    }
     public static int putearnedcredit(String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int cellno)
     {
//         System.out.println("earned credits cellno"+cellno);
          CellStyle style = workbook.createCellStyle();
          style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
        
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        
        
         //extracted form pdf
          Pattern earn=Pattern.compile("ARNED :" );
          Matcher m=earn.matcher(text);  
          if(m.find())
          {
          String totalcreditsearned =text.substring(m.end(),m.end()+3);
         totalcreditsearned= totalcreditsearned.replaceAll("[^0-9]", "");
          
            Row row = sheet.getRow(rownum-1); 
//              System.out.println("earned credits cellno1"+cellno);
            Cell cell = row.createCell(cellno);  
            
            cellno++;
//              System.out.println("earned credits cellno2"+cellno);
            cell=row.createCell(cellno);
 cell.setCellStyle(style);
//   System.out.println("earned credits cellno3"+cellno);
 if(isStringInt(totalcreditsearned)){
              int val=Integer.parseInt(totalcreditsearned);
               cell.setCellValue(val); 
            }
            else{
                
                cell.setCellValue(totalcreditsearned); 
            }
            
//              System.out.println("earned credits cellno4"+cellno);
            
             return Integer.parseInt(totalcreditsearned);
          }
          return -1;
     }
    public static void init(int year,int sem ,String text,int cellnum,int rownum,XSSFSheet sheet,XSSFWorkbook workbook,String student) {
        
        int k=0,n=0;
        getSubjectCode(year,student,sheet,workbook);
             switch (year) {
             case 2:
                 k=6;n=12;
                 break;
             case 3:
                 k=8; n=16;
                 break;
             case 4:
                 k=8;n=15;
                 break;
         }
//         if(sem==0){
//             n=k;
//         }
            sem1end=k;sem2end=n; 
             System.out.println("parameter sem="+sem1end+"  " +sem2end);
             for(int i=0;i<n;i++){
                  System.out.println("subn =i="+i+" "+subn[i]);
                 if(subn[i]==1){
                     if(i<k){
                          sem1end++;
                     }
                     sem2end++;
                 }
             }
             System.out.println("parameter sem="+sem1end+"  " +sem2end);
                sem1end=(sem1end*5)+3;
                sem2end=(sem2end*5)+6;
             
                 System.out.println("parameter sem="+sem1end+"  " +sem2end);
                
    	
    	CellStyle style = workbook.createCellStyle(); 
    	CellStyle style1 = workbook.createCellStyle();
        style1.setAlignment(CellStyle.ALIGN_CENTER);
        style1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        Font font = workbook.createFont();  
        Font font1 = workbook.createFont();  
        font1.setFontHeightInPoints((short)20);  
        font1.setFontName("Courier New");  
        
        font.setBold(true);  
        font1.setBold(true); 
        // Applying font to the style  
        style1.setFont(font1); 
       // style.setFont(font); 
        
        style1.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
        
        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
    	Row row0 = sheet.createRow(1); 
    	row0.setHeight((short)500);
    	cellnum=2;
    	Cell cell1 = row0.createCell(cellnum++); 
        
        cell1.setCellValue("\t\t   RESULT"); 
        
        style1.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
        style1.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style1.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
        style1.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style1.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
        style1.setTopBorderColor(IndexedColors.BLACK.getIndex());

        cell1.setCellStyle(style1);
        for(int i=0;i<sem2end+3;i++) {
        cell1 = row0.createCell(cellnum++); 
        cell1.setCellStyle(style1);
        }
//        ---------------------------------
        style1.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
        style1.setRightBorderColor(IndexedColors.BLACK.getIndex());  
        
        cell1.setCellStyle(style1);
        CellStyle style2 = workbook.createCellStyle();
        style2.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
        style2.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style2.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
        style2.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style2.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
        style2.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style2.setFont(font); 
        style2.setAlignment(CellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style2.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());  
        
        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
        sheet.addMergedRegion(new CellRangeAddress(1,1,2,sem2end+5));  
        sheet.addMergedRegion(new CellRangeAddress(3,3,3,sem1end));  
        sheet.addMergedRegion(new CellRangeAddress(3,3,sem1end+2,sem2end));  
        sheet.addMergedRegion(new CellRangeAddress(5,6,0,0)); 
	sheet.addMergedRegion(new CellRangeAddress(5,6,1,1)); 
	

		rownum=3;
        row0 = sheet.createRow(rownum++); 
        cellnum=0;
        row0.setHeight((short)500);
        cellnum=3;
        cell1 = row0.createCell(cellnum++); 
        
        cell1.setCellValue("\t\tSEM 1"); 
        style1.setFont(font1); 
        cell1.setCellStyle(style1);
        for(int i=0;i<sem1end-3;i++) {
            cell1 = row0.createCell(cellnum++); 
            cell1.setCellStyle(style1);
            }
//        if(sem2){
        style1.setFont(font1); 
        cell1.setCellStyle(style1);
        cell1 = row0.createCell(cellnum++); 
        cell1 = row0.createCell(cellnum++);
        cell1.setCellValue("\t\tSEM 2");
        cell1.setCellStyle(style1);
        for(int i=0;i<sem2end-sem1end-2;i++) {
            cell1 = row0.createCell(cellnum++); 
            cell1.setCellStyle(style1);
            } 
//    }
        rownum=5;
    	row0 = sheet.createRow(rownum++); 
    	 row0.setHeight((short)400);
    	cellnum=0;
        Cell cell0 = row0.createCell(cellnum++); 
        cell0.setCellValue("  Sr No.  "); 
        cell0.setCellStyle(style2);
        cell0 = row0.createCell(cellnum++); 
        cell0.setCellValue("Seat no"); 
        cell0.setCellStyle(style2);
        cell0=row0.createCell(cellnum++);
        cell0.setCellValue("Subject Code");
        cell0.setCellStyle(style2);
        
        // Setting Foreground Color  
        style = workbook.createCellStyle(); 
        
        style.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setFont(font); 
       // style.setFillForegroundColor(IndexedColors.BLUE.getIndex());  
       // style.setFillPattern(CellStyle.SOLID_FOREGROUND);

        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
   
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());  
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
    //    style.setVerticalAlignment(VerticalAlignment.CENTER);
        int cm=3;
         
//                System.out.println("sun= 11 "+sub[11]);
//                System.out.println("sun=12 "+sub[12]); 
       // System.out.println("sub l "+sub.length); 
        for (int i=1; i<=n; i++){
            
             if(sub[i-1]==null)
                break;
//            System.out.println("sun="+(i)+" "+subn[i-1]); 
//        	System.out.println("sun="+(i)+" "+sub[i-1]);
                
                
        	cell0=row0.createCell(cellnum++);
        	cell0.setCellValue(sub[i-1]);
       	    cell0.setCellStyle(style);
        	 if(subn[i-1]==0) {
                     if(i==k+1) {
        			 cm=cm+2;
        		 }
        		 sheet.addMergedRegion(new CellRangeAddress(5,5,cm,cm+4)); 
        		 cm=cm+5;
        		 for (int j=0; j<4; j++){
             		
            		 cell0=row0.createCell(cellnum++);
            		 cell0.setCellStyle(style);
            	 }
        		 
        	 }
        	 else {
        		 if(i==k+1) {
        			 cm=cm+2;
        		 }
        		 sheet.addMergedRegion(new CellRangeAddress(5,5,cm,cm+9)); 
        		 cm=cm+10;
        		 for (int j=0; j<9; j++){
             		
            		 cell0=row0.createCell(cellnum++);
            		 cell0.setCellStyle(style);
            	 }
        		
        	 }
        	
        	 cell0.setCellStyle(style);
        	 if(i==k||i==n) {
        		 cell0=row0.createCell(cellnum++);

        	        style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        	        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        	    	cell0.setCellValue("Sem "+(i/6));
        	        cell0.setCellStyle(style); 
        	        cell0=row0.createCell(cellnum++);
        	    	cell0.setCellValue("        ");
        	   
        	        
        	 }
        	
     	}
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
        
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        if(year==4){
             cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("Earned Credits");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("SGPA");
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("SGPA (PDF)");
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("Result");
        cell0.setCellStyle(style);
               cell0=row0.createCell(cellnum++);
               cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("FE SGPA");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("SE SGPA");
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("TE SGPA");
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("CGPA");
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("Final Result");
        cell0.setCellStyle(style);
        }
        else
        {     cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("Earned Credits");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("SGPA");
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("SGPA (PDF)");
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    	cell0.setCellValue("Result");
        cell0.setCellStyle(style);
        }


//        ===============================================================================
        
        row0 = sheet.createRow(rownum++); 
        row0.setHeight((short)400);
        cell0=row0.createCell(2);
        cell0.setCellValue("Name");
        cell0.setCellStyle(style); 
        cellnum=3;
        style = workbook.createCellStyle(); 
        style.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
          style.setAlignment(CellStyle.ALIGN_CENTER); 
          style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        style.setFillForegroundColor(IndexedColors.TAN.getIndex());  
        
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(font); 
        cell0=row0.createCell(0);
        
   	    cell0.setCellStyle(style);
   	 cell0=row0.createCell(1);
	    cell0.setCellStyle(style);
   		for (int j=0; j<n; j++){
   		 cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("  TOT  ");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("  Crd  ");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("Grd");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("Grdpts");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("Crdpts");
        cell0.setCellStyle(style);
        if(subn[j]==1){
        	cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("TW");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("Crd");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("Grd");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("Grdpts");
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
   		cell0.setCellValue("Crdpts");
        cell0.setCellStyle(style);
        }
        
        if(j==k-1||j==n-1) {
        	CellStyle style3 = workbook.createCellStyle(); 
            style3.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
            style3.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
            style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
            style3.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
            style3.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
            style3.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
            style3.setTopBorderColor(IndexedColors.BLACK.getIndex());
            style3.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
            style3.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
              style3.setAlignment(CellStyle.ALIGN_CENTER); 
              style3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        	style3.setFillForegroundColor(IndexedColors.BLUE.getIndex());  
            style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
        	cell0=row0.createCell(cellnum++);
       		cell0.setCellValue("  TOTCrdpts  ");
            cell0.setCellStyle(style3);
            cell0=row0.createCell(cellnum++);
       		cell0.setCellValue("      ");
           
        
        }
        style.setFillForegroundColor(IndexedColors.TAN.getIndex());  
        
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
   	 }
//style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
        
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell0=row0.createCell(cellnum++);
    	
        cell0.setCellStyle(style); 
        cell0=row0.createCell(cellnum++);
    	
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    
        cell0.setCellStyle(style);
        cell0=row0.createCell(cellnum++);
    	
        cell0.setCellStyle(style);
   	 
    }
    public static void getSubjectCode(int y,String text, XSSFSheet sheet,XSSFWorkbook workbook)
    {
         String year="";
          switch (y) {
             case 2:
                 year= "20";
                 break;
             case 3:
                 year = "30";
                 break;
             case 4:
                 year = "40";
                 break;
         }
    	  Pattern subject=Pattern.compile(year+"[0-9]{4}" );
          Matcher m=subject.matcher(text); 
          Pattern sem2p=Pattern.compile("SEM.:2" );
          Matcher m2=subject.matcher(text);  
//          if(m2.find())
//             sem2=true;
//          else
//             sem2=false; 
          String prev="";
          int i=0;
          if(m.find())
              prev=text.substring(m.start(),m.end());
          sub[i]=prev;
          while(m.find())
          {  
            String temp=text.substring(m.start(),m.end());
            if(temp.equals(prev))
                subn[i]=1;
            else
            {
                i++;
                sub[i]=temp;
                prev=temp;
            }
                 
               
          } 
    } 
    public static void putSubjectData(XSSFWorkbook workbook,String[] studentdata, XSSFSheet sheet, Row row,int cellno,int y,int sem1,int sem2m)
    { 
    	 CellStyle style = workbook.createCellStyle();
          //style.setFillForegroundColor(IndexedColors.TURQUOISE .getIndex());  
//         System.out.println("putsub data cellno"+cellno);
        //style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
        style.setAlignment(CellStyle.ALIGN_CENTER); 
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
         int sem1position=0, sem2position=0;
         switch(y)
         {
             case 2:
             {
                 sem1position=58;
                 sem2position=110;
                 break;
             }
             case 3:
             {
                 sem1position=53;
                 sem2position=105;
                 break;
             }
             case 4:
             {
                 sem1position=53;
                 sem2position=100;
                 break;
             }
                 
         }
//          System.out.println("earned credits stu data len"+studentdata.length);
    	for(int i=0; studentdata[i]!=null && i<studentdata.length;i++)
    	{
//              System.out.println("cellno="+cellno);
            if(cellno==sem1position) 
            { 
                CellStyle style3 = workbook.createCellStyle(); 
                style3.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
                style3.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
                style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
                style3.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
                style3.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
                style3.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
                style3.setTopBorderColor(IndexedColors.BLACK.getIndex());
                style3.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
                style3.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
                style3.setAlignment(CellStyle.ALIGN_CENTER); 
                style3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
                style3.setFillForegroundColor(IndexedColors.BLUE.getIndex());  
                style3.setFillPattern(CellStyle.SOLID_FOREGROUND);

                Cell cell=row.createCell(cellno++);
                cell.setCellStyle(style3);                        
                cell.setCellValue(sem1); 
                cellno++;  
            }
            
            Cell cell = row.createCell(cellno++); 
//            System.out.println("FF here");
            if(isStringInt(studentdata[i])){
              int val=Integer.parseInt(studentdata[i]);
                cell.setCellValue(val);
            }
            else{
                
                cell.setCellValue(studentdata[i]);
            }
            
            cell.setCellStyle(style); 
//            if(sem2){
            if(cellno==sem2position)
            {
                CellStyle style3 = workbook.createCellStyle(); 
                style3.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
                style3.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
                style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
                style3.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);  
                style3.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
                style3.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);  
                style3.setTopBorderColor(IndexedColors.BLACK.getIndex());
                style3.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);  
                style3.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
                style3.setAlignment(CellStyle.ALIGN_CENTER); 
                style3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
                style3.setFillForegroundColor(IndexedColors.BLUE.getIndex());  
                style3.setFillPattern(CellStyle.SOLID_FOREGROUND);

                cell=row.createCell(cellno++);
                cell.setCellStyle(style3);
                cell.setCellValue(sem2m); 
                cellno++;
            }
//            }
    	}
//        System.out.println("cellno="+sem2);
//        System.out.println("cellno="+cellno);
//        if(sem2){
            
//        }
//        else{
//            System.out.println("cellno="+cellno);
             Cell cell=row.createCell(cellno++);
//                cell.setCellStyle(style3);                        
                cell.setCellValue(sem1); 
                cell=row.createCell(cellno++);
//        }
        
    } 
    public static String[] getSubjectData(String text) throws InterruptedException
    {  
    	
        int size=0;
        for(int i=0;i<sub.length;i++)
        {
            if(subn[i]==1)
                size+=10;
            else if(sub[i]!=null)
                size+=5;
        }
    	String studentdata[]=new String[size+11]; 
    	int j=0;
    	for(int k=0;k<sub.length;k++)
        {
            if(sub[k]==null)
                break;
    	    Pattern subjectcode=Pattern.compile(sub[k]);            
            Matcher sub=subjectcode.matcher(text);      
            int end;
            while(sub.find())
            { 
                int  start=sub.start()+startfactor;
                if((int)text.charAt(sub.end())>64 && ((int)text.charAt(sub.end())<70) )
                {
                     
                   end=sub.start()+endfactor+2;
                }
                else
               end=sub.start()+endfactor+1;
                String marks=text.substring(start,end);
                String arr[]=marks.split(" ");  
                 for(int i=0;i<arr.length;i++)
                 { 
                    if(arr[i].length()>=1 && arr[i]!=null)
                    {
                        studentdata[j]=arr[i];
                        j++;
                    }
                 }  
            } 
    	}  
	return studentdata;
    }
    public static int  getSeatNoAndName(int year ,String text, XSSFWorkbook workbook, XSSFSheet sheet,int rownum,int no) throws InterruptedException
    {
//    	  System.out.println("getseatNo start");
    	  Pattern seatNo;
         switch (year) {
             case 2:
                 seatNo = Pattern.compile("S"+"[0-9]{9}" );
                 break;
             case 3:
                 seatNo = Pattern.compile("T"+"[0-9]{9}" );
                 break;
             default:
                 seatNo = Pattern.compile("B"+"[0-9]{9}" );
                 break;
         }

          Matcher m=seatNo.matcher(text); 
          int cellnum=0;
//          Row row0 = sheet.createRow(rownum++); 
//          Cell cell0 = row0.createCell(cellnum++); 
//          cell0.setCellValue("Seat no"); 
//          cell0=row0.createCell(cellnum++);
//          cell0.setCellValue("Name"); 
          while(m.find())
          {  
             // cellnum=0;
        	CellStyle style = workbook.createCellStyle();  
        	style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
        	style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
                style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
                style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
                style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
                style.setTopBorderColor(IndexedColors.BLACK.getIndex());
                style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
                style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
                style.setAlignment(CellStyle.ALIGN_CENTER); 
                style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); 
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                style.setAlignment(CellStyle.ALIGN_CENTER);  
                style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); 
              Row row = sheet.createRow(rownum++);     
              Cell cell = row.createCell(cellnum++); 
              cell.setCellStyle(style);
              cell.setCellValue(no); 
              cell = row.createCell(cellnum++); 
              cell.setCellValue(text.substring(m.start(), m.end())); 
              
              String name=getName(text.substring(m.end()+2,m.end()+50));
              
              System.out.println(no+") "+ name);
              cell.setCellStyle(style);
              cell=row.createCell(cellnum++);
              cell.setCellValue(name); 
               style = workbook.createCellStyle();
               style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());  
//             
             style.setFillPattern(CellStyle.SOLID_FOREGROUND);
             style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
             style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); 
             style.setFillPattern(CellStyle.SOLID_FOREGROUND);
             style.setBorderBottom(XSSFCellStyle.BORDER_THIN);  
             style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
             style.setBorderTop(XSSFCellStyle.BORDER_THIN);  
             style.setTopBorderColor(IndexedColors.BLACK.getIndex());
             style.setBorderRight(XSSFCellStyle.BORDER_THIN);  
             style.setRightBorderColor(IndexedColors.BLACK.getIndex()); 
             style.setAlignment(CellStyle.ALIGN_CENTER); 
             style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
             cell.setCellStyle(style); 
            
              String studentdata[]=getSubjectData(text);
               
              int sem_marks[]=new int[2];
//              System.out.println("sem mark start");
              sem_marks=calculate_semmarks(studentdata,year);
             
//              System.out.println(sem1+" "+ sem2);
            //  Thread.sleep(1);
//            System.out.println("put sub data start");
              putSubjectData(workbook,studentdata,sheet,row,cellnum,year,sem_marks[0],sem_marks[1]);
                 numof_students++;
          } 
//          System.out.println("getseatNo end");
          return rownum;
//          System.out.println("getseatNo end");
    }
    public static int[] calculate_semmarks(String studentdata[],int y )
    {
         int count=0;
         int arr[]=new int[2];
         int semdivider=0;
         switch(y){
             case 2:
                 semdivider=11;
                 break;
             case 3:
                 semdivider=10;
                 break;
             case 4:
                 semdivider=10;
                 break;
         }  
              for(int i=1;i<studentdata.length;i++)
              { 
                  if(studentdata[i-1]!=null && i%5==0)
                  {
                      if(count<semdivider)
                      {
                          arr[0]+=Integer.parseInt(studentdata[i-1]);
                          count++;
                      }
                      else
                      {
                            arr[1]+=Integer.parseInt(studentdata[i-1]);
                            count++;
                      }
                  }
              }
              return arr;
    }
    public static String getName(String text)
    {
        int start=0;
    	if(text.charAt(start)==' ')
    	start++;
        
    	String name="";  
    	while(true)
    	{
    		if(text.charAt(start)==' ' && text.charAt(start+1)==' ' || start==47)
    			break;
    		name+=text.charAt(start);
    		start++;    				
    	}    
    	return name;
    }
    
    public static void getSubTop(XSSFWorkbook workbook, XSSFSheet sheet,int year,int n1)
   {
       int k=0,n=0;
       switch (year) {
             case 2:
                 k=6;n=12;
                 break;
             case 3:
                 k=8; n=16;
                 break;
             case 4:
                 k=8;n=15;
                 break;
         }
//         if(sem2){
//             
//         }
//         else{
//             n=k;
//         }
       int rownum=n1+9;
//       String c[]={"D","N","X","AH","AR","BB","BI","BS","BX","CH","CR","CW"};

       int x=4,ntop=1,itop=0;
//       System.out.println("here "+n1);
       Row row=sheet.createRow(rownum); 
       Row row1=sheet.createRow(rownum+1); 
       Row row2=sheet.createRow(rownum+2); 
//       System.out.println("herecell"+k);
       for(int i=0;i<n;i++){
           Cell cell = row.createCell(x-1);
       cell.setCellFormula("MAX("+excelColName(x)+"8:"+excelColName(x)+(n1+7)+")");
       cell.setCellType(Cell.CELL_TYPE_FORMULA);
       cell = row.createCell(x+4);
       cell.setCellFormula("MAX("+excelColName(x+5)+"8:"+excelColName(x+5)+(n1+7)+")");
       cell.setCellType(Cell.CELL_TYPE_FORMULA);
        ///Count top
       cell = row1.createCell(x-1);
       cell.setCellFormula("COUNTIF("+excelColName(x)+"8:"+excelColName(x)+(n1+7)+","+excelColName(x)+(n1+10)+")");
       cell.setCellType(Cell.CELL_TYPE_FORMULA);
//       ntop=(int)row1.getCell(x-1).getNumericCellValue();
//       System.out.println("ntop="+ntop);
//       for(int j=0;j<ntop;j++){
//           row2=sheet.createRow(rownum+2+j); 
//           Cell cell1 = row2.createCell(x-2);
//         cell1.setCellFormula("INDEX(A8:DG"+(n1+7)+",MATCH("+excelColName(x)+""+(n1+10)+","+excelColName(x)+(8+itop)+":"+excelColName(x)+(n1+7)+",0),1)");
//         itop=(int)row1.getCell(x-1).getNumericCellValue();
//         System.out.println("itop="+itop);
           Cell cell1 = row2.createCell(x-1);
         cell1.setCellFormula("INDEX(A8:DG"+(n1+7)+",MATCH("+excelColName(x)+""+(n1+10)+","+excelColName(x)+(8+itop)+":"+excelColName(x)+(n1+7)+",0),3)");
       cell1.setCellType(Cell.CELL_TYPE_FORMULA);
//       }
       
       System.out.println(excelColName(x));
      if(subn[i]==1){
          
       x+=10;
       }else 
       {
           x+=5;
       }
       if(i==(k-1)){
          x+=2;
      } 
    
   }
       x=111;
       for(int i=0;i<5;i++){
           Cell cell = row.createCell(x-1);
       cell.setCellFormula("MAX("+excelColName(x)+"8:"+excelColName(x)+(n1+7)+")");
       System.out.println(excelColName(x));
       cell.setCellType(Cell.CELL_TYPE_FORMULA); 
       x++;
       }
      
       
       
   }
    public static void getSubCount(XSSFWorkbook workbook, XSSFSheet sheet,int year,int n1)
   {
       int k=0,n=0;
       switch (year) {
             case 2:
                 k=6;n=12;
                 break;
             case 3:
                 k=8; n=16;
                 break;
             case 4:
                 k=8;n=15;
                 break;
         }
//         if(sem2){
//             
//         }
//         else{
//             n=k;
//         }
       int rownum=n1+15;
        int x=3;
       System.out.println("herecount"+n1);
       Row row=sheet.createRow(rownum++); 
       Row row1=sheet.createRow(rownum+1); 
       Row row2=sheet.createRow(rownum+2); 
       Row row3=sheet.createRow(rownum+3); 
       Row row4=sheet.createRow(rownum+4); 
       Row row5=sheet.createRow(rownum+5); 
       Cell cell = row.createCell(x-1);
       cell.setCellValue("Count");
       Cell cell1 = row1.createCell(x-1);
       cell1.setCellValue("Total:");
        Cell cell2 = row2.createCell(x-1);
       cell2.setCellValue("FF:");
        Cell cell3 = row3.createCell(x-1);
       cell3.setCellValue("Pass:");
        Cell cell4 = row4.createCell(x-1);
       cell4.setCellValue("Failure %");
        Cell cell5 = row5.createCell(x-1);
       cell5.setCellValue("Passing %");
          System.out.println("herecell"+k);
          x++;
       for(int i=0;i<n*2-1;i++){
         
       //count tot
       cell = row1.createCell(x-1);
       cell.setCellFormula("COUNTA("+excelColName(x)+"8:"+excelColName(x)+(n1+7)+")");
       cell.setCellType(Cell.CELL_TYPE_FORMULA);
       //count ff
       cell = row2.createCell(x-1);
       cell.setCellFormula("COUNTIF("+excelColName(x)+"8:"+excelColName(x)+(n1+7)+",\"FF\")");
       cell.setCellType(Cell.CELL_TYPE_FORMULA);
       //pass
       cell = row3.createCell(x-1);
       cell.setCellFormula(""+excelColName(x)+(rownum+2)+"-"+(excelColName(x)+(rownum+3)));
       cell.setCellType(Cell.CELL_TYPE_FORMULA);
       //pass%
       cell = row4.createCell(x-1);
       cell.setCellFormula(""+excelColName(x)+(rownum+3)+"/"+(excelColName(x)+(rownum+2))+"*100");
       cell.setCellType(Cell.CELL_TYPE_FORMULA);
       //fail%
       cell = row5.createCell(x-1);
       cell.setCellValue(25);
       cell.setCellFormula("100"+"-"+excelColName(x)+(rownum+5));
//       cell.setCellType(Cell.CELL_TYPE_FORMULA);
       System.out.println(excelColName(x));
       
//      if(subn[i]==1){
//          
//       x+=10;
//       }else 
//       {
//           x+=5;
//       }
       if((n==16&&i==10)||(n==12&&i==10)||(n==15&&i==9)){
          x+=2;
      }
       else{
           x+=5;
       }
    
   }

   
   }
     public static void putCongdata(XSSFWorkbook workbook, XSSFSheet sheet,XSSFSheet sheet1,int year,int n1){
         
         int k=0,n=0;
          String s[]={};
       switch (year) {
             case 2:
                 k=6;n=12;
                 String s1[]={
         "SIGNALS AND SYSTEMS","ELECTRONIC DEVICES AND CIRCUITS ","ELECTRICAL CIRCUITS AND MACHINES","DATA STRUCTURES AND ALGORITHM","DIGITAL ELECTRONICS",
         "ELECTRONICS MEASURING INSTRUMENTS AND TOOLS","INTEGRATED CIRCUITS","CONTROL SYSTEMS ","ANALOG COMMUNICATION","OBJECT ORIENTED PROGRAMMING ","EMPLOYABILITY SKILL DEVELOPMENT"
         ,"ENGINEERING MATHEMATICS III "};
                 s=s1;
                 break;
             case 3:
                 k=8; n=16;
                 String s2[]={
         "304181","304182","304183","304184","304185","SIGN. PROC. & COMM. LAB","MICROCON. & MECHATRONICS LAB","ELECTRONICS SYSTEM DESIGN","Power Electronics","Information Theory,Coding and Communication\n" +
"Networks","Business Management","Advanced Processes","System Programming And Operationg System","Power And ITCT Lab","Advanced Processes and System Programming Lab","Employability Skills and Mini Project"
         };
                 s=s2;
                 break;
             case 4:
                 k=8;n=15;
                 s=sub;
                
                 break;
         }
         
         
         int rownum=n1+22,cellno=4;
         int rownum1=3,cellno1=3;
         Row row1=sheet.getRow(rownum); 
//         Row row2=sheet.getRow(6);
          System.out.println("getrow succesfully executed"+" sem2="+sem2);
          Row row=sheet1.createRow(rownum1++); 
           System.out.println("create succesfully executed"+" sem2="+sem2);
          Cell cell=row.createCell(cellno1++);
          cell.setCellValue("SR No.");
          cell=row.createCell(cellno1++);
          cell.setCellValue("Subject (SEM 1)");
          cell=row.createCell(cellno1++);
          cell.setCellValue(" TH%  .");
          cell=row.createCell(cellno1++);
          cell.setCellValue("TW/PR/OR %");
           System.out.println("create 2 succesfully executed"+" sem2="+sem2);
          row=sheet1.createRow(rownum1++);
          
           System.out.println("for succesfully executed"+" sem2="+sem2);
           double th=0,tw=0;
           int sr=1;
//           Cell cell2=row2.createCell(3);
//           String s1=cell2.getStringCellValue();
          for(int j=0;j<n;j++){
              cellno1=3;
              if(j==k){
                  row=sheet1.createRow(rownum1++);
                  row=sheet1.createRow(rownum1++);
            cell=row.createCell(cellno1++);
//           cell.setCellValue(j+1);
          cell=row.createCell(cellno1++);
          cell.setCellValue("SUBJECT(SEM-II)");
          cellno+=2;
          row=sheet1.createRow(rownum1++);
//          continue;
              }
              cellno1=3;
               cell=row.createCell(cellno1++);
           cell.setCellValue(sr);
           sr++;
          cell=row.createCell(cellno1++);
          cell.setCellValue(s[j]);
           cell=row.createCell(cellno1++);
            System.out.println("create cell succesfully executed"+" sem2="+rownum);

            System.out.println(" succesfully executed"+" cellno="+cellno);
            if(subn[sr-2]==1){
                
          cell.setCellFormula("ROUND('student Marksheet'!"+excelColName(cellno)+rownum+",2)");
           cellno+=5;
            }
            
         
//            while(s1.equals("TW")==false){
//        cellno++;
//        cell2=row2.getCell(cellno);
//        s1=cell2.getStringCellValue();
//    }
          cell=row.createCell(cellno1++);
          System.out.println(" succesfully executed"+" cellno="+cellno);
          cell.setCellFormula("ROUND('student Marksheet'!"+excelColName(cellno)+rownum+",2)");
//          ROUND('student Marksheet'!S349,2)
//          while(s1.equals("TOT")==false){
//        cellno++;
//        cell2=row2.getCell(cellno);
//        s1=cell2.getStringCellValue();
//    }
          cellno+=5;
 System.out.println(" succesfully executed"+" cellno="+cellno);
          row=sheet1.createRow(rownum1++);
          }
        
         
         
         
     }
      public static void getOverAlldata(XSSFWorkbook workbook, XSSFSheet sheet,XSSFSheet sheet1,int year,int n1){
      
      int cellno=sem2end+8,rownum=n1+23;
      int res =sem2end+5;
      Row row=sheet.createRow(rownum++); 
       Cell cell = row.createCell(cellno++);
      cell.setCellValue("Total");
      cell = row.createCell(cellno);
      cell.setCellValue(n1);
      row=sheet.createRow(rownum++); 
      cellno=sem2end+8;
      cell = row.createCell(cellno++);
      cell.setCellValue("FD");
        cell = row.createCell(cellno++);
       cell.setCellFormula("COUNTIF("+excelColName(res)+"8:"+excelColName(res)+(n1+7)+",\"FD\")");
       cell = row.createCell(cellno++);
       cell.setCellFormula("ROUND("+excelColName(cellno-1)+""+(rownum)+"/"+excelColName(cellno-1)+(rownum-1)+"*100,3)");
      row=sheet.createRow(rownum++); 
      cellno=sem2end+8;
      cell = row.createCell(cellno++);
      cell.setCellValue("FC");
        cell = row.createCell(cellno++);
       cell.setCellFormula("COUNTIF("+excelColName(res)+"8:"+excelColName(res)+(n1+7)+",\"FC\")");
       cell = row.createCell(cellno++);
       cell.setCellFormula(""+excelColName(cellno-1)+""+(rownum)+"/"+excelColName(cellno-1)+(rownum-2)+"*100");
      row=sheet.createRow(rownum++); 
      cellno=sem2end+8;
      cell = row.createCell(cellno++);
      cell.setCellValue("FAIL");
      
        cell = row.createCell(cellno++);
       cell.setCellFormula("COUNTIF("+excelColName(res)+"8:"+excelColName(res)+(n1+7)+",\"FF\")");
       cell = row.createCell(cellno++);
       cell.setCellFormula(""+excelColName(cellno-1)+""+(rownum)+"/"+excelColName(cellno-1)+(rownum-3)+"*100");
      row=sheet.createRow(rownum++); 
      cellno=sem2end+8;
      cell = row.createCell(cellno++);
      cell.setCellValue("ALL");
        cell = row.createCell(cellno++);
      cell.setCellFormula("SUM("+excelColName(cellno)+(rownum-3)+":"+excelColName(cellno)+(rownum-1)+")");
//        cell.setCellValue("T");
        cell = row.createCell(cellno++);
       cell.setCellFormula("ROUND("+excelColName(cellno-1)+""+(rownum)+"/"+excelColName(cellno-1)+(rownum-4)+"*100,3)");
      }
      public static void putOverAlldata(XSSFWorkbook workbook, XSSFSheet sheet,XSSFSheet sheet1,int year,int n1){
      
          int rownum=n1+24,cellno=sem2end+10;
         int rownum1=3,cellno1=9;
         Row row1=sheet.getRow(rownum);
          Row row=sheet1.getRow(rownum1++); 
//           System.out.println();
          Cell cell=row.createCell(cellno1);
          cell.setCellValue("TOTAL NO OF STUDENT APPEARED");
          cell=row.createCell(cellno1+1);
          cell.setCellValue(n1);
          row=sheet1.getRow(rownum1++); 
          cell=row.createCell(cellno1);
          cell.setCellValue("RESULT");
          cell=row.createCell(cellno1+1);
           cell.setCellValue("NO OF STUDENTS (%)");
           
           row=sheet1.getRow(rownum1++); 
           cell=row.createCell(cellno1);
          cell.setCellValue("ALL CLEAR:");
          cell=row.createCell(cellno1+1);
//           cell.setCellValue("0");
System.out.println(" succesfully executed put"+" cellno="+cellno);
//String ss1="CONCATENATE('student Marksheet'!"+excelColName(cellno)+rownum+",'student Marksheet'!"+excelColName(cellno+1)+rownum+")";
//System.out.println("put ss1= "+ss1);

  cell.setCellFormula("CONCATENATE('student Marksheet'!"+excelColName(cellno)+(rownum+4)+",\" (\",'student Marksheet'!"+excelColName(cellno+1)+(rownum+4)+",\")\")");
           System.out.println(" formula succesfully executed"+" cellno="+cellno);
           row=sheet1.getRow(rownum1++); 
           cell=row.createCell(cellno1);
          cell.setCellValue("DISTINCTION ( > 7.75 SGPA)");
          cell=row.createCell(cellno1+1);
//           cell.setCellValue("0");
rownum++;
cell.setCellFormula("CONCATENATE('student Marksheet'!"+excelColName(cellno)+rownum+",\" (\",'student Marksheet'!"+excelColName(cellno+1)+rownum+",\")\")");
           
           row=sheet1.getRow(rownum1++); 
           cell=row.createCell(cellno1);
          cell.setCellValue("FIRST CLASS (  6.75  TO  7.74 SGPA )");
          cell=row.createCell(cellno1+1);
//           cell.setCellValue("0");
rownum++;
cell.setCellFormula("CONCATENATE('student Marksheet'!"+excelColName(cellno)+rownum+",\" (\",'student Marksheet'!"+excelColName(cellno+1)+rownum+",\")\") ");
//                  "CONCATENATE('SE MAY 2020      '!      DK294             , " (" ,'SE MAY 2020      '!      DL294               , ") ") "
           row=sheet1.getRow(rownum1++); 
           cell=row.createCell(cellno1);
          cell.setCellValue("HIGH.SECOND CLASS ( 6.25  TO  6.74  SGPA )");
          cell=row.createCell(cellno1+1);
//           cell.setCellValue("0");
rownum++;
cell.setCellFormula("CONCATENATE('student Marksheet'!"+excelColName(cellno)+rownum+",\" (\",'student Marksheet'!"+excelColName(cellno+1)+rownum+",\")\")");
           
           row=sheet1.getRow(rownum1++); 
           cell=row.createCell(cellno1);
          cell.setCellValue("SECOND CLASS  ( 5.5  TO  6.24  SGPA )");
          cell=row.createCell(cellno1+1);
           cell.setCellValue("0(0)");
rownum++;
//cell.setCellFormula("CONCATENATE('student Marksheet'!"+excelColName(cellno)+rownum+",\" (\",'student Marksheet'!"+excelColName(cellno+1)+rownum+",\")\")");
      
      
      
      }
      
        public static void getTopperdata(XSSFWorkbook workbook, XSSFSheet sheet,XSSFSheet sheet1,int year,int n1){
      
          int rownum=n1+24,cellno=sem2end+10;
         int rownum1=5,cellno1=13;
         Row row1=sheet.getRow(rownum);
          Row row=sheet1.getRow(rownum1++); 
          Cell cell1,cell2;
          double max=0,no=0;
         double top[]=new double[3];
         String[] tn=new String[3];
         tn[0]=" ";tn[1]=" ";tn[2]=" ";
         String name;
         for(int i=0;i<n1;i++){
//             System.out.println("i="+i);
           row1=sheet.getRow(i+7);
            cell1=row1.getCell(sem2end+3);
           try{
            no=Double.parseDouble(cell1.getStringCellValue());}
           catch(Exception e){
               continue;
           }
            if(no>top[0]){
                top[2]=top[1];
                top[1]=top[0];
                top[0]=no;
            }else if(no>top[1]&&no!=top[0]){
                top[2]=top[1];
                top[1]=no;
                
            }
            else if(no>top[2]&&no!=top[1]&&no!=top[0]){
                top[2]=no;
            }
         }
         for(int i=0;i<n1;i++){
//             System.out.println("ii="+i);
           row1=sheet.getRow(i+7);
            cell1=row1.getCell(sem2end+3);
            try{
            no=Double.parseDouble(cell1.getStringCellValue());}
           catch(Exception e){
               continue;
           }
            
         if(no==top[0]){
            cell2=row1.getCell(2);
            name=cell2.getStringCellValue();
            tn[0]=tn[0]+","+name;
             
            }else if(no==top[1]){
                cell2=row1.getCell(2);
            name=cell2.getStringCellValue();
            tn[1]=tn[1]+","+name; 
                
            }
            else if(no==top[2]){
               cell2=row1.getCell(2);
            name=cell2.getStringCellValue();
            tn[2]=tn[2]+","+name;
            }
         
         
         
         }
         System.out.println(top[0]+" || "+top[1]+" || "+top[2]);
         System.out.println(tn[0]+" || "+tn[1]+" || "+tn[2]);
         System.out.println("for done");
         String []tnsp;
         System.out.println("tnsp done");
            tnsp= tn[0].split(",");
            System.out.println("split done");
            cell1=row.createCell(cellno1++);
            System.out.println("cell done");
            cell1.setCellValue(" S.N. ");
           cell1=row.createCell(cellno1++);
            cell1.setCellValue("NAME OF STUDENT");
            cell1=row.createCell(cellno1++);
            cell1.setCellValue("SGPA");
            cellno1=13;
        row=sheet1.getRow(rownum1++);
        System.out.println("row done");
          cell1=row.createCell(cellno1++);
            cell1.setCellValue("1st");
           cell1=row.createCell(cellno1++);
            cell1.setCellValue(tnsp[0]);
            System.out.println("tnsp val done");
            cell1=row.createCell(cellno1);
            cell1.setCellValue(top[0]);
            for(int i=1;i<tnsp.length;i++){
            row=sheet1.getRow(rownum1++);
            cell1=row.createCell(cellno1-1);
            cell1.setCellValue(tnsp[i]);
            }
            
            System.out.println("1st done");
            tnsp=tn[1].split(",");
            System.out.println("2nd split done");
            cellno1=13;
        row=sheet1.getRow(rownum1++);
          cell1=row.createCell(cellno1++);
            cell1.setCellValue("2nd");
           cell1=row.createCell(cellno1++);
            cell1.setCellValue(tnsp[0]);
            cell1=row.createCell(cellno1);
            cell1.setCellValue(top[1]);
            for(int i=1;i<tnsp.length;i++){
            row=sheet1.getRow(rownum1++);
            cell1=row.createCell(cellno1-1);
            cell1.setCellValue(tnsp[i]);
            }
         tnsp= tn[2].split(",");
         cellno1=13;
        row=sheet1.getRow(rownum1++);
          cell1=row.createCell(cellno1++);
            cell1.setCellValue("3rd");
           cell1=row.createCell(cellno1++);
            cell1.setCellValue(tnsp[0]);
            cell1=row.createCell(cellno1);
            cell1.setCellValue(top[2]);
            for(int i=1;i<tnsp.length;i++){
            row=sheet1.getRow(rownum1++);
            cell1=row.createCell(cellno1-1);
            cell1.setCellValue(tnsp[i]);
            }
        
        }
   public static boolean isStringInt(String s)
{
    try
    {
        Integer.parseInt(s);
        return true;
    } catch (NumberFormatException ex)
    {
        try
    {
        Float.parseFloat(s);
        return true;
    } catch (NumberFormatException ex1)
    {
        return false;
    }
        
    }
    
} 
   private static String excelColName(int columnNumber)
    {
        // To store result (Excel column name)
        StringBuilder columnName = new StringBuilder();
 
        while (columnNumber > 0) {
            // Find remainder
            int rem = columnNumber % 26;
 
            // If remainder is 0, then a
            // 'Z' must be there in output
            if (rem == 0) {
                columnName.append("Z");
                columnNumber = (columnNumber / 26) - 1;
            }
            else // If remainder is non-zero
            {
                columnName.append((char)((rem - 1) + 'A'));
                columnNumber = columnNumber / 26;
            }
        }
 
        // Reverse the string and print result
        String s=columnName.reverse().toString();
        return s;
    }
    
}

   
