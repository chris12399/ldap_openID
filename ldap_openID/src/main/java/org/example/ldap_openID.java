package org.example;

import java.io.*;

import com.sun.xml.internal.bind.v2.model.core.ID;
import jdk.nashorn.internal.runtime.arrays.IteratorAction;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.commons.io.FilenameUtils;

import java.util.*;

import javax.naming.AuthenticationException;
import javax.naming.Context;
import javax.naming.NamingEnumeration;
import javax.naming.NamingException;
import javax.naming.directory.*;

public class ldap_openID {
    String displayName_global = null;
    String uid_global = null;
    String st_global = null;
    String employeeNumber_global = null;


    public void createFile(File outputFile, Map file_data) {
//        String excelFilePath = "java_result.xls";
        Workbook workbook = null;
        Row row = null;
        String nameCell = null;
        String IDCell = null;
        int columnCount =14;
        Iterator<String> it1 = file_data.keySet().iterator();
        HashSet<String> hashSet = new HashSet<>();

        try{

/*
            FileInputStream fis = new FileInputStream(inputFile);
            String ext  = FilenameUtils.getExtension(inputFile.toString());
            if(ext.equalsIgnoreCase("xlsx")) {
                workbook = new XSSFWorkbook(fis);
            }else if(ext.equalsIgnoreCase("xls")){
                workbook = new HSSFWorkbook(fis);
            }
            Sheet sheet = workbook.getSheetAt(0);

            for(int i = 1; i <= sheet.getLastRowNum(); i++) {
                row = sheet.getRow(i);
                nameCell = row.getCell(5).toString(); //姓名
                IDCell = row.getCell(8).toString();   //身分證

                hashSet.add(nameCell+"_"+IDCell);
            }
            int n=0;
            Iterator<String> it2 = hashSet.iterator();
            while (it1.hasNext()) {
                while(it2.hasNext()){
                    if(it1.next().equals(it2)){
                        n++;
                    }
                }
            }
*/



            OutputStreamWriter osw = new OutputStreamWriter(new FileOutputStream(outputFile),"UTF-8");
            BufferedWriter bw = new BufferedWriter(osw);

            String [] titles = new String[] {"student", "openID"};
            for(String title : titles){
                bw.write(title);
                bw.write(",");
            }
            bw.write("\r\n");
            for(Object key : file_data.keySet()) {
                bw.write(key.toString());
                bw.write(",");
                bw.write(file_data.get(key).toString());
                bw.write("\r\n");
            }
            bw.flush();
            bw.close();
            osw.close();

        }catch (Exception e) {
            e.printStackTrace();
        }

    }


    public Map readFile(File inputFile, Map LDAP_data) throws IOException {

        LinkedHashMap<String, String> file_data= new LinkedHashMap<String, String>();
        FileInputStream fis = new FileInputStream(inputFile);
//        OutputStreamWriter osw = new OutputStreamWriter(new FileOutputStream(outputFile),"UTF-8");
//        BufferedWriter bw = new BufferedWriter(osw);
        Workbook workbook = null;
        Row row = null;
        String school_cell = null;
        String name_cell = null;
        String ID_cell = null;
//        LinkedHashMap<String, String> nameAndID = new LinkedHashMap<String, String>();


//        ArrayList<String> temps = new ArrayList<String>();
//        Iterator<String> it1 = LDAP_data.keySet().iterator();
        int z =0;
        try {
            String ext  = FilenameUtils.getExtension(inputFile.toString());
            if(ext.equalsIgnoreCase("xlsx")) {
                workbook = new XSSFWorkbook(fis);
            }else if(ext.equalsIgnoreCase("xls")){
                workbook = new HSSFWorkbook(fis);
            }

            Sheet sheet = workbook.getSheetAt(0);

            for(int i = 1; i <= sheet.getLastRowNum(); i++) {
                row = sheet.getRow(i);
                school_cell = row.getCell(1).toString(); //學校名稱
                name_cell = row.getCell(5).toString(); //姓名
                ID_cell = row.getCell(8).toString();   //身分證

//                nameAndID.put(name_cell, ID_cell);
//                temps.add(name_cell+"_"+ID_cell);
//                String test = LDAP_data.get(name_cell).toString();
//                if(temps.get(i).equals(it1.next())) {
//                    z++;
//                }

                file_data.put(name_cell+"_"+ID_cell, "無配對");

            }
            //System.out.println("this is file_data size " + file_data.size());
        } catch (Exception e) {
            e.printStackTrace();
        }

        int cont = 0;
        for (Object obj1 : LDAP_data.keySet()) {
            for (Object obj2 : file_data.keySet()) {
                if (obj1.equals(obj2)) {
                    cont++;
                    file_data.replace(obj2.toString(), LDAP_data.get(obj2).toString());
                    break;
                }
            }
        }
//        System.out.println(file_data + "\n");
        System.out.println(cont);
        return file_data;
    }



    public Map connect(String url, String username, String password) {
        DirContext ctx = null;
        Hashtable<String, String> HashEnv = new Hashtable<String, String >();
        HashEnv.put(Context.SECURITY_AUTHENTICATION, "simple");
        HashEnv.put(Context.SECURITY_PRINCIPAL, username);
        HashEnv.put(Context.SECURITY_CREDENTIALS, password);
        HashEnv.put(Context.INITIAL_CONTEXT_FACTORY, "com.sun.jndi.ldap.LdapCtxFactory");
        HashEnv.put("com.sun.jndi.ldap.connect.timeout", "3000");//連線超時設定為3秒
        HashEnv.put(Context.PROVIDER_URL, url);

        LinkedHashMap<String, String> LDAP_data = new LinkedHashMap<String, String>();
//        LinkedHashMap<String, String> nameAndID = new LinkedHashMap<String, String>();
        NamingEnumeration<?> namingEnum = null;
        try {
            ctx = new InitialDirContext(HashEnv);
            System.out.println("驗證成功");

            String dnBase = "ou=school,dc=ldap,dc=kl,dc=edu,dc=tw";

            namingEnum = ctx.search(dnBase, "l=student", getSearchControls());
            while (namingEnum.hasMore()) {
                SearchResult result = (SearchResult) namingEnum.next();
                Attributes attrs = result.getAttributes();
                String displayName = getAttributeValue(attrs.get("displayName")); //姓名
                String uid = getAttributeValue(attrs.get("uid")); //openID
                String st = getAttributeValue(attrs.get("st")); //學校
                String employeeNumber = getAttributeValue(attrs.get("employeeNumber")); //身分證

//                nameAndID.put(displayName, employeeNumber);
                LDAP_data.put(displayName+"_"+employeeNumber, uid);

            }
            //System.out.println(LDAP_data + "\n");
            //System.out.println(LDAP_data.size());

        } catch (AuthenticationException e) {
            System.out.println("驗證失敗");
            e.printStackTrace();
        } catch (javax.naming.CommunicationException e) {
            System.out.println("連線失敗");
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(null != ctx){
                try {
                    ctx.close();
                    ctx = null;
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return LDAP_data;
    }

    private  String getAttributeValue(Attribute attr) throws NamingException {
        String val = null;
        if(attr != null) {
            Object obj = attr.get();
            if(obj instanceof byte[]) {
                val = new String((byte[]) obj);
            }else {
                val = obj.toString();
            }
        }
        return val;
    }

    private SearchControls getSearchControls() {
        SearchControls searchControls = new SearchControls();
        searchControls.setSearchScope(SearchControls.SUBTREE_SCOPE);
        searchControls.setTimeLimit(50000);
        return searchControls;
    }


//**************************************************************************

    public static void main(String[] args) throws IOException {

        ldap_openID ldID = new ldap_openID();
        File inputFile = new File("/home/user/圖片/1024_LDAP/export (8).xls");
        File outputFile = new File("/home/user/圖片/1024_LDAP/result.xls");


        Map set1 = ldID.connect("ldap://210.240.1.242:389",
                "cn=root,dc=ldap,dc=kl,dc=edu,dc=tw","bp45/4@n4zj6fu4");
        Map set2 = ldID.readFile(inputFile, set1);

        ldID.createFile(outputFile, set2);
        System.out.println("*********Hello world! this is ldap openID compare code***********");

    }
}