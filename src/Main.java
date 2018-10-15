
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.nio.cs.Surrogate.Generator;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author BHT
 */
public class Main {
    public static int ID=14000;
    public static int REVISION_ID=53000;
    private void writePropMappingFromExcel(String input, String mapping, String ontologyClass, String output) throws IOException {
        // generate string
        SimpleDateFormat sdf= new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
        FileInputStream fis = new FileInputStream(new File(input));
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        StringBuilder sb= new StringBuilder();
        String str="<page><title>Mapping vi:"+mapping+"</title><ns>382</ns><id>"+ID+"</id><revision><id>"+REVISION_ID+"</id><timestamp>"+sdf.format(new Date())+"</timestamp><text>{{TemplateMapping \n" +
"| mapToClass = "+ontologyClass+"\n" +
"| mappings = \n";
        sb.append(str);
        for (Row row : sheet) {
            if(row.getCell(0)!=null) {
                if(row.getCell(2)!=null && row.getCell(2).toString()!="") {
                    str="{{ PropertyMapping | templateProperty = "+row.getCell(0).toString()+" | ontologyProperty = "+row.getCell(2).toString()+" }}\n";
                    sb.append(str);
                }
            }
        }
        sb.append("}}</text></revision></page>");
        // write to file
        try {
            PrintWriter pw = new PrintWriter(new FileOutputStream(new File(output),true)); 
            pw.print(sb);
            pw.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Generator.class.getName()).log(Level.SEVERE, null, ex);
        }
        ID++;
        REVISION_ID++;
    }

    private void writeClassMappingFromExcel(String input, String output) throws IOException {
        SimpleDateFormat sdf= new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
        FileInputStream fis = new FileInputStream(new File(input));
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        StringBuilder sb= new StringBuilder();
        for (Row row : sheet) {
            if(row.getRowNum()!=0&&row.getCell(0)!=null&&row.getCell(0).toString()!=""&&row.getCell(1)!=null&& row.getCell(1).toString()!=""&&row.getCell(2)!=null&& row.getCell(2).toString()!="") {
                String str="<page><title>Mapping vi:"+row.getCell(1).toString().substring(0,row.getCell(1).toString().length()-1)+"</title><ns>382</ns><id>"+ID+"</id><revision><id>"+REVISION_ID+"</id><timestamp>"+sdf.format(new Date())+"</timestamp><text>{{TemplateMapping \n" +
                        "| mapToClass = "+row.getCell(2).toString()+"\n" +
                        "| mappings = \n}}</text></revision></page>";
                sb.append(str);
                ID++;
                REVISION_ID++;
            }
        }
        // write to file
        try {
            PrintWriter pw = new PrintWriter(new FileOutputStream(new File(output),true));
            pw.print(sb);
            pw.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Generator.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void initMappingFile(String output) throws IOException{
        StringBuilder sb= new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><mediawiki xmlns=\"http://www.mediawiki.org/xml/export-0.8/\">");
        try {
            PrintWriter pw = new PrintWriter(new FileOutputStream(new File(output),true));
            pw.print(sb);
            pw.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Generator.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void closeMappingFile(String output) throws IOException {
        StringBuilder sb= new StringBuilder();
        sb.append("</mediawiki>");
        try {
            PrintWriter pw = new PrintWriter(new FileOutputStream(new File(output),true));
            pw.print(sb);
            pw.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Generator.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void main(String[] args) throws IOException {
        String output="Mapping_vi.xml";
        Main m= new Main();
        m.initMappingFile(output);
        m.writePropMappingFromExcel("Thông_tin_phim_result.xlsx", "Thông tin phim", "Film", output);
        m.writePropMappingFromExcel("Thông_tin_núi_result.xlsx", "Thông tin núi", "Mountain", output);
        m.writePropMappingFromExcel("Thông_tin_chùa_result.xlsx", "Thông tin chùa", "Temple", output);
        m.writePropMappingFromExcel("Thông_tin_hồ_result.xlsx", "Thông tin hồ", "Lake", output);
        m.writePropMappingFromExcel("Thông_tin_nhạc_sĩ_result.xlsx", "Thông tin nhạc sĩ", "MusicalArtist", output);
        m.writePropMappingFromExcel("Thông_tin_sách_result.xlsx", "Thông tin sách", "Book", output);
        m.writePropMappingFromExcel("Thông_tin_cầu_result.xlsx", "Thông tin cầu", "Bridge", output);
        m.writePropMappingFromExcel("Bảng_phân_loại_result.xlsx", "Bảng phân loại", "Categories", output);
        m.writePropMappingFromExcel("Thông_tin_nhà_ga_result.xlsx", "Thông tin nhà ga", "Station", output);
        m.writePropMappingFromExcel("Thông_tin_sân_bay_result.xlsx", "Thông tin sân bay", "Airport", output);
        m.writePropMappingFromExcel("Thông_tin_món_ăn_result.xlsx", "Thông tin món ăn", "Food", output);
        m.writePropMappingFromExcel("Thông_tin_bài_hát_result.xlsx", "Thông tin bài hát", "Song", output);
        m.writePropMappingFromExcel("Thông_tin_hành_tinh_result.xlsx", "Thông tin hành tinh", "Planet", output);
        m.writePropMappingFromExcel("Thông_tin_nhà_văn_result.xlsx", "Thông tin nhà văn", "Writer", output);
        m.writePropMappingFromExcel("Thông_tin_album_nhạc_result.xlsx", "Thông tin album nhạc", "Album", output);
        m.writePropMappingFromExcel("Thông_tin_dân_tộc_result.xlsx", "Thông tin dân tộc", "EthnicGroup", output);
        m.writePropMappingFromExcel("Thông_tin_đĩa_đơn_result.xlsx", "Thông tin đĩa đơn", "Single", output);
        m.writePropMappingFromExcel("Thông_tin_ngày_lễ_result.xlsx", "Thông tin ngày lễ", "Holiday", output);
        m.writePropMappingFromExcel("Thông_tin_nghệ_sĩ_result.xlsx", "Thông tin nghệ sĩ", "Artist", output);
        m.writePropMappingFromExcel("Thông_tin_hóa_chất_result.xlsx", "Thông tin hóa chất", "ChemicalElement", output);
        m.writePropMappingFromExcel("Thông_tin_khu_dân_cư_result.xlsx", "Thông tin khu dân cư", "Settlement", output);
        m.writePropMappingFromExcel("Thông_tin_nhân_vật_result.xlsx", "Thông tin nhân vật", "Person", output);
        m.writePropMappingFromExcel("Thông_tin_diễn_viên_result.xlsx", "Thông tin diễn viên", "Actor", output);
        m.writePropMappingFromExcel("Thông_tin_viên_chức_result.xlsx", "Thông tin viên chức", "OfficeHolder", output);
        m.writePropMappingFromExcel("Thông_tin_chiến_tranh_result1.xlsx", "Thông tin chiến tranh", "MilitaryConflict", output);
        m.writePropMappingFromExcel("Thông_tin_phần_mềm_result.xlsx", "Thông tin phần mềm", "Software", output);
        m.writePropMappingFromExcel("Thông_tin_nhà_khoa_học_result1.xlsx", "Thông tin nhà khoa học", "Scientist", output);
        m.writePropMappingFromExcel("Thông_tin_truyền_hình_result.xlsx", "Thông tin truyền hình", "Televisonshow", output);
        m.writePropMappingFromExcel("Tiểu_sử_quân_nhân_result.xlsx", "Tiểu sử quân nhân", "MilitaryPerson", output);
        m.writePropMappingFromExcel("Hộp_thông_tin_vũ_khí_result.xlsx", "Hộp thông tin vũ khí", "Weapon", output);
        m.writePropMappingFromExcel("Thông_tin_trường_học_result.xlsx", "Thông tin trường học", "EducationalInstitution", output);
        m.writePropMappingFromExcel("Tóm_tắt_về_công_ty_result.xlsx", "Tóm tắt về công ty", "Company", output);
        m.writePropMappingFromExcel("Thông_tin_người_mẫu_result.xlsx", "Thông tin người mẫu", "Model", output);
        m.writePropMappingFromExcel("Hộp_thông_tin_triết_gia_result.xlsx", "Hộp thông tin triết gia", "Philosopher", output);
        m.writePropMappingFromExcel("Thông_tin_đảng_phái_chính_trị_result.xlsx", "Thông tin đảng phái chính trị", "PoliticalParty", output);
        m.writePropMappingFromExcel("Thông_tin_điện_thoại_di_động_result.xlsx", "Thông tin điện thoại di động", "MobilePhone", output);
        m.writePropMappingFromExcel("Thông_tin_Di_sản_thế_giới_result.xlsx", "Thông tin Di sản thế giới", "WorldHeritageSite", output);
        m.writePropMappingFromExcel("Thông_tin_đơn_vị_hành_chính_Việt_Nam_result.xlsx", "Thông tin đơn vị hành chính Việt Nam", "AdminitrativeRegion", output);
        m.writePropMappingFromExcel("Thông_tin_đơn_vị_quân_sự_result.xlsx", "Thông tin đơn vị quân sự", "MilitaryUnit", output);
        m.writePropMappingFromExcel("Thông_tin_khu_vực_bảo_tồn_result.xlsx", "Thông tin khu vực bảo tồn", "ProtectedArea", output);
        m.writePropMappingFromExcel("Thông_tin_nguyên_tố_hóa_học_result.xlsx", "Thông tin nguyên tố hóa học", "ChemicalElement", output);
        m.writePropMappingFromExcel("Thông_tin_nhân_vật_hoàng_gia_result.xlsx", "Thông tin nhân vật hoàng gia", "Royalty", output);
        m.writePropMappingFromExcel("Thông_tin_nhân_vật_hư_cấu_result.xlsx", "Thông tin nhân vật hư cấu", "FictionalCharacter", output);
        m.writePropMappingFromExcel("Thông_tin_tiểu_sử_bóng_đá_result.xlsx", "Thông tin tiểu sử bóng đá", "SoccerPlayer", output);
        m.writePropMappingFromExcel("Thông_tin_trò_chơi_điện_tử_result.xlsx", "Thông tin trò chơi điện tử", "Game", output);
        m.writePropMappingFromExcel("Thông_tin_về_sân_vận_động_result.xlsx", "Thông tin về sân vận động", "Stadium", output);
        m.writePropMappingFromExcel("Tóm_tắt_về_ngôn_ngữ_result.xlsx", "Tóm tắt về ngôn ngữ", "Language", output);


        m.writeClassMappingFromExcel("Kết-quả-ánh-xạ-lớp.xlsx",output);
        m.closeMappingFile(output);
    }
}
