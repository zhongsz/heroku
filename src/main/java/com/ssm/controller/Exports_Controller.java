package com.ssm.controller;

import java.io.OutputStream;
import java.util.ArrayList;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.ssm.bean.LAP_Name;

/**
 * 
 * @author Administrator
 *
 */
@SuppressWarnings("deprecation")
@Controller
public class Exports_Controller {

	@ResponseBody
	@RequestMapping("/LAP_Exports")
	public void export(HttpServletRequest request, HttpServletResponse response) {

		/*******************************************************************************/

		/*****************************创建数据测试*********************************************/
		LAP_Name LAP1 = new LAP_Name("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "QR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", null, null, null, null, null, null, null, "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", null, null, null, "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EYS", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", null, "FY", "FZ");
		LAP_Name LAP2 = new LAP_Name("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "QR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", null, null, null, null, null, null, null, "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", null, null, null, "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EYS", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", null, "FY", "FZ");
		LAP_Name LAP3 = new LAP_Name("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "QR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", null, null, null, null, null, null, null, "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", null, null, null, "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EYS", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", null, "FY", "FZ");
		LAP_Name LAP4 = new LAP_Name("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "QR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", null, null, null, null, null, null, null, "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", null, null, null, "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EYS", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", null, "FY", "FZ");
		LAP_Name LAP5 = new LAP_Name("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "QR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", null, null, null, null, null, null, null, "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", null, null, null, "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EYS", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", null, "FY", "FZ");
		
		ArrayList<LAP_Name> list = new ArrayList<LAP_Name>();
		list.add(LAP1);
		list.add(LAP2);
		list.add(LAP3);
		list.add(LAP4);
		list.add(LAP5);
		
		
		/*****************************创建数据测试*********************************************/
		
		HSSFWorkbook wb = Exports_Controller.export_text(list);

		try {
			// 定义导出文件的名称，看不懂的同学可以先行了解一下文件下载
			String fileName = new String("personRelation.xls".getBytes("UTF-8"), "ISO-8859-1");
			response.setContentType("application/vnd.ms-excel");
			response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
			OutputStream os = response.getOutputStream();

			// 将工作薄写入到输出流中
			wb.write(os);
			os.close();
		} catch (Exception e) {
			e.getStackTrace();
		}

		/*******************************************************************************/
	}

	/**
	 * 创建标题行样式
	 * 
	 * @param wb
	 * @return
	 */
	/********************************************************************************/
	private static HSSFWorkbook export_text(ArrayList<LAP_Name> list) {
		try {
			// 创建工作薄对象
			HSSFWorkbook wb = new HSSFWorkbook();
			
			
			
			//创建标题行样式
            CellStyle headStyle = headStyle(wb);

			// 创建sheet页并且命名
			HSSFSheet sheet_1 = wb.createSheet("Apttus_Proposal__Proposal");

			//设置表的默认列宽
            sheet_1.setDefaultColumnWidth(20);
            
            //创建标题行
            //获取插入标头的行数
            HSSFRow row = sheet_1.createRow(0);
            HSSFRow row1 = sheet_1.createRow(1);
            
     /*************************************第一行标头*****************************************************/       
            HSSFCell NP_Flag = row.createCell(0);           //创建标题行第一列
            NP_Flag.setCellValue("NP Flag");                        //第一列内容
            NP_Flag.setCellStyle(headStyle);

			HSSFCell BE_Flag = row.createCell(1);
			BE_Flag.setCellValue("BE Flag");
			BE_Flag.setCellStyle(headStyle);
			
			HSSFCell TS_Flag = row.createCell(2);
			TS_Flag.setCellValue("TS Flag");
			TS_Flag.setCellStyle(headStyle);
			
			HSSFCell HL_Item = row.createCell(3);
			HL_Item.setCellValue("HL Item");
			HL_Item.setCellStyle(headStyle);
			
			HSSFCell Item_No = row.createCell(4);
			Item_No.setCellValue("Item No");
			Item_No.setCellStyle(headStyle);
			
			HSSFCell Part_Number = row.createCell(5);
			Part_Number.setCellValue("Part Number");
			Part_Number.setCellStyle(headStyle);
			
			HSSFCell Description = row.createCell(6);
			Description.setCellValue("Description");
			Description.setCellStyle(headStyle);
			
			HSSFCell Standard_Price = row.createCell(7);
			Standard_Price.setCellValue("Standard Price");
			Standard_Price.setCellStyle(headStyle);
			
			HSSFCell List_Price = row.createCell(8);
			List_Price.setCellValue("List Price");
			List_Price.setCellStyle(headStyle);
			
			HSSFCell Discount_from_list = row.createCell(9);
			Discount_from_list.setCellValue("Discount from list");
			Discount_from_list.setCellStyle(headStyle);
			
			HSSFCell Requested = row.createCell(10);
			Requested.setCellValue("Requested");
			Requested.setCellStyle(headStyle);
			
			HSSFCell Entitled = row.createCell(12);
			Entitled.setCellValue("Entitled");
			Entitled.setCellStyle(headStyle);
			
			HSSFCell Adjusted_CQ_Cost = row.createCell(15);
			Adjusted_CQ_Cost.setCellValue("Adjusted CQ Cost");
			Adjusted_CQ_Cost.setCellStyle(headStyle);
			
			HSSFCell Quantity = row.createCell(16);
			Quantity.setCellValue("Quantity");
			Quantity.setCellStyle(headStyle);
			
			HSSFCell Total_Qty = row.createCell(20);
			Total_Qty.setCellValue("Total Qty");
			Total_Qty.setCellStyle(headStyle);
			
			HSSFCell Contract_Start = row.createCell(21);
			Contract_Start.setCellValue("Contract Start");
			Contract_Start.setCellStyle(headStyle);
			
			HSSFCell Contract_End = row.createCell(22);
			Contract_End.setCellValue("Contract End");
			Contract_End.setCellStyle(headStyle);
			
			HSSFCell D_Price = row.createCell(23);
			D_Price.setCellValue("D Price");
			D_Price.setCellStyle(headStyle);
			
			HSSFCell C_Price = row.createCell(24);
			C_Price.setCellValue("C Price");
			C_Price.setCellStyle(headStyle);
			
			HSSFCell B_Price = row.createCell(25);
			B_Price.setCellValue("B Price");
			B_Price.setCellStyle(headStyle);
			
			HSSFCell YBMD = row.createCell(26);
			YBMD.setCellValue("(YBMD)");
			YBMD.setCellStyle(headStyle);
			
			HSSFCell YBMC = row.createCell(27);
			YBMC.setCellValue("(YBMC)");
			YBMC.setCellStyle(headStyle);
			
			HSSFCell YBMB = row.createCell(28);
			YBMB.setCellValue("(YBMB)");
			YBMB.setCellStyle(headStyle);
			
			HSSFCell Warranty = row.createCell(29);
			Warranty.setCellValue("Warranty");
			Warranty.setCellStyle(headStyle);
			
			HSSFCell Sales_Feasibility_Comments = row.createCell(31);
			Sales_Feasibility_Comments.setCellValue("Sales Feasibility Comments");
			Sales_Feasibility_Comments.setCellStyle(headStyle);
			
			HSSFCell Feasibility_approver_Comments = row.createCell(32);
			Feasibility_approver_Comments.setCellValue("Feasibility approver Comments");
			Feasibility_approver_Comments.setCellStyle(headStyle);
			
			HSSFCell Commissionable_Cost = row.createCell(33);
			Commissionable_Cost.setCellValue("Commissionable Cost");
			Commissionable_Cost.setCellStyle(headStyle);
			
			HSSFCell Feasibility_Cost_adjustment = row.createCell(37);
			Feasibility_Cost_adjustment.setCellValue("Feasibility  Cost adjustment");
			Feasibility_Cost_adjustment.setCellStyle(headStyle);
			
			HSSFCell Cost_Adjustment = row.createCell(38);
			Cost_Adjustment.setCellValue("Cost Adjustment");
			Cost_Adjustment.setCellStyle(headStyle);
			
			HSSFCell Adjusted_Comm_Cost = row.createCell(43);
			Adjusted_Comm_Cost.setCellValue("Adjusted Comm Cost");
			Adjusted_Comm_Cost.setCellStyle(headStyle);
			
			HSSFCell Base_or_Profit_Flag = row.createCell(47);
			Base_or_Profit_Flag.setCellValue("Base or Profit Flag");
			Base_or_Profit_Flag.setCellStyle(headStyle);
			
			HSSFCell A_Price = row.createCell(48);
			A_Price.setCellValue("A Price");
			A_Price.setCellStyle(headStyle);
			
			HSSFCell YBMA = row.createCell(49);
			YBMA.setCellValue("(YBMA)");
			YBMA.setCellStyle(headStyle);
			
			HSSFCell Revenue_Reduction = row.createCell(50);
			Revenue_Reduction.setCellValue("Revenue Reduction %");
			Revenue_Reduction.setCellStyle(headStyle);
			
			HSSFCell Revenue_Reduction_simulation = row.createCell(51);
			Revenue_Reduction_simulation.setCellValue("Revenue Reduction % simulation");
			Revenue_Reduction_simulation.setCellStyle(headStyle);
			
			HSSFCell Sev_Pac = row.createCell(52);
			Sev_Pac.setCellValue("Sev Pac");
			Sev_Pac.setCellStyle(headStyle);
			
			HSSFCell M1_Cost = row.createCell(53);
			M1_Cost.setCellValue("M1 Cost");
			M1_Cost.setCellStyle(headStyle);
			
			HSSFCell M2_Cost = row.createCell(54);
			M2_Cost.setCellValue("M2 Cost");
			M2_Cost.setCellStyle(headStyle);
			
			HSSFCell M3_Cost = row.createCell(55);
			M3_Cost.setCellValue("M3 Cost");
			M3_Cost.setCellStyle(headStyle);
			
			HSSFCell M4_Cost = row.createCell(56);
			M4_Cost.setCellValue("M4 Cost");
			M4_Cost.setCellStyle(headStyle);
			
			HSSFCell M5_Cost = row.createCell(57);
			M5_Cost.setCellValue("M5 Cost");
			M5_Cost.setCellStyle(headStyle);
			
			HSSFCell M6_Cost = row.createCell(58);
			M6_Cost.setCellValue("M6 Cost");
			M6_Cost.setCellStyle(headStyle);
			
			HSSFCell M7_Cost = row.createCell(59);
			M7_Cost.setCellValue("M7 Cost");
			M7_Cost.setCellStyle(headStyle);
			
			
			HSSFCell M8_Cost = row.createCell(60);
			M8_Cost.setCellValue("M8 Cost");
			M8_Cost.setCellStyle(headStyle);
			
			HSSFCell M9_Cost = row.createCell(61);
			M9_Cost.setCellValue("M9 Cost");
			M9_Cost.setCellStyle(headStyle);
			
			HSSFCell M10_Cost = row.createCell(62);
			M10_Cost.setCellValue("M10 Cost");
			M10_Cost.setCellStyle(headStyle);
			
			HSSFCell M11_Cost = row.createCell(63);
			M11_Cost.setCellValue("M11 Cost");
			M11_Cost.setCellStyle(headStyle);
			
			HSSFCell M12_Cost = row.createCell(64);
			M12_Cost.setCellValue("M12 Cost");
			M12_Cost.setCellStyle(headStyle);
			
			HSSFCell PH_value = row.createCell(65);
			PH_value.setCellValue("PH value");
			PH_value.setCellStyle(headStyle);
			
			HSSFCell PH_description = row.createCell(66);
			PH_description.setCellValue("PH description");
			PH_description.setCellStyle(headStyle);
			
			HSSFCell Margin = row.createCell(67);
			Margin.setCellValue("Margin %");
			Margin.setCellStyle(headStyle);
			
			HSSFCell MarginS = row.createCell(72);
			MarginS.setCellValue("Margin $");
			MarginS.setCellStyle(headStyle);
			
			HSSFCell TOTAL_GROSS_REVENUE = row.createCell(76);
			TOTAL_GROSS_REVENUE.setCellValue("TOTAL GROSS REVENUE");
			TOTAL_GROSS_REVENUE.setCellStyle(headStyle);
			
			HSSFCell TOTAL_NET_REVENUE = row.createCell(80);
			TOTAL_NET_REVENUE.setCellValue("TOTAL NET REVENUE");
			TOTAL_NET_REVENUE.setCellStyle(headStyle);
			
			HSSFCell TOTAL_Comm_COST_Inc_adjustments = row.createCell(84);
			TOTAL_Comm_COST_Inc_adjustments.setCellValue("TOTAL Comm COST (Inc adjustments)");
			TOTAL_Comm_COST_Inc_adjustments.setCellStyle(headStyle);
			
			HSSFCell TOTAL_MARGIN_after_revenue_reduction = row.createCell(88);
			TOTAL_MARGIN_after_revenue_reduction.setCellValue("TOTAL MARGIN after revenue reduction");
			TOTAL_MARGIN_after_revenue_reduction.setCellStyle(headStyle);
			
			HSSFCell Margin_after_rev_reduction = row.createCell(92);
			Margin_after_rev_reduction.setCellValue("Margin% after rev reduction");
			Margin_after_rev_reduction.setCellStyle(headStyle);
			
			HSSFCell Adjusted_Comm_Cost_CS = row.createCell(96);
			Adjusted_Comm_Cost_CS.setCellValue("Adjusted Comm Cost");
			Adjusted_Comm_Cost_CS.setCellStyle(headStyle);
			
			HSSFCell Term_Fx_Cost = row.createCell(97);
			Term_Fx_Cost.setCellValue("Term Fx Cost");
			Term_Fx_Cost.setCellStyle(headStyle);
			
			HSSFCell Cost_adjustment_CU = row.createCell(98);
			Cost_adjustment_CU.setCellValue("Cost adjustment %");
			Cost_adjustment_CU.setCellStyle(headStyle);
			
			HSSFCell Payment_terms_Cost = row.createCell(99);
			Payment_terms_Cost.setCellValue("Payment terms Cost");
			Payment_terms_Cost.setCellStyle(headStyle);
			
			HSSFCell Term_Cost_Adjustment = row.createCell(100);
			Term_Cost_Adjustment.setCellValue("Term Cost Adjustment");
			Term_Cost_Adjustment.setCellStyle(headStyle);
			
			HSSFCell Cost_Adjustment_Note = row.createCell(101);
			Cost_Adjustment_Note.setCellValue("Cost Adjustment Note");
			Cost_Adjustment_Note.setCellStyle(headStyle);
			
			HSSFCell Total_Term_Cost = row.createCell(102);
			Total_Term_Cost.setCellValue("Total Term Cost");
			Total_Term_Cost.setCellStyle(headStyle);
			
			HSSFCell Total_Term_Margin = row.createCell(103);
			Total_Term_Margin.setCellValue("Total Term Margin");
			Total_Term_Margin.setCellStyle(headStyle);
		
//			HSSFCell DA = row.createCell(104);
//			DA.setCellValue("");
//			DA.setCellStyle(headStyle);
//			
//			HSSFCell DB = row.createCell(105);
//			DB.setCellValue("");
//			DB.setCellStyle(headStyle);
//			
//			HSSFCell DC = row.createCell(106);
//			DC.setCellValue("");
//			DC.setCellStyle(headStyle);
//			
//			HSSFCell DD = row.createCell(107);
//			DD.setCellValue("");
//			DD.setCellStyle(headStyle);
//			
//			HSSFCell DE = row.createCell(108);
//			DE.setCellValue("");
//			DE.setCellStyle(headStyle);
//			
//			HSSFCell DF = row.createCell(109);
//			DF.setCellValue("");
//			DF.setCellStyle(headStyle);
//			
//			HSSFCell DG = row.createCell(110);
//			DG.setCellValue("");
//			DG.setCellStyle(headStyle);
			
			HSSFCell Cycle = row.createCell(111);
			Cycle.setCellValue("Cycle");
			Cycle.setCellStyle(headStyle);
			
			HSSFCell Standard_Price_DI = row.createCell(112);
			Standard_Price_DI.setCellValue("Standard Price");
			Standard_Price_DI.setCellStyle(headStyle);
			
			HSSFCell List_Price_DJ = row.createCell(113);
			List_Price_DJ.setCellValue("List Price");
			List_Price_DJ.setCellStyle(headStyle);
			
			HSSFCell Adjusted_CQ_Cost_DK = row.createCell(114);
			Adjusted_CQ_Cost_DK.setCellValue("Adjusted CQ Cost");
			Adjusted_CQ_Cost_DK.setCellStyle(headStyle);
			
			HSSFCell Requested_DL_DN = row.createCell(115);
			Requested_DL_DN.setCellValue("Requested");
			Requested_DL_DN.setCellStyle(headStyle);
			
			HSSFCell Entitled_DO_DR = row.createCell(118);
			Entitled_DO_DR.setCellValue("Entitled");
			Entitled_DO_DR.setCellStyle(headStyle);
			
			HSSFCell PAB_Price = row.createCell(122);
			PAB_Price.setCellValue("PAB Price");
			PAB_Price.setCellStyle(headStyle);
			
			HSSFCell Active_Price = row.createCell(125);
			Active_Price.setCellValue("Active Price");
			Active_Price.setCellStyle(headStyle);
			
			HSSFCell Solution_ID = row.createCell(128);
			Solution_ID.setCellValue("Solution ID");
			Solution_ID.setCellStyle(headStyle);
			
			HSSFCell List_Price_Adjustment = row.createCell(129);
			List_Price_Adjustment.setCellValue("List Price Adjustment");
			List_Price_Adjustment.setCellStyle(headStyle);
			
			HSSFCell List_Price_Adjustment_Note = row.createCell(130);
			List_Price_Adjustment_Note.setCellValue("List Price Adjustment Note");
			List_Price_Adjustment_Note.setCellStyle(headStyle);
			
			HSSFCell Auto_Renew = row.createCell(131);
			Auto_Renew.setCellValue("Auto Renew");
			Auto_Renew.setCellStyle(headStyle);
		
			/*************************后加***************************/
			HSSFCell Guidance_for_PWT = row.createCell(135);
			Guidance_for_PWT.setCellValue("Guidance for PWT");
			Guidance_for_PWT.setCellStyle(headStyle);
			
		//******
			HSSFCell Line_Number = row.createCell(145);
			Line_Number.setCellValue("Line Number");
			Line_Number.setCellStyle(headStyle);
			
			HSSFCell Item_Sequence = row.createCell(146);
			Item_Sequence.setCellValue("Item Sequence");
			Item_Sequence.setCellStyle(headStyle);
			
			HSSFCell Is_Configuration_Instruction = row.createCell(147);
			Is_Configuration_Instruction.setCellValue("Is Configuration Instruction");
			Is_Configuration_Instruction.setCellStyle(headStyle);
			
			HSSFCell Is_Contracted = row.createCell(148);
			Is_Contracted.setCellValue("Is Contracted");
			Is_Contracted.setCellStyle(headStyle);
			
			HSSFCell ExtendedCost = row.createCell(149);
			ExtendedCost.setCellValue("ExtendedCost");
			ExtendedCost.setCellStyle(headStyle);
			
		//***********
			HSSFCell Is_Top_Seller = row.createCell(150);
			Is_Top_Seller.setCellValue("Is Top Seller");
			Is_Top_Seller.setCellStyle(headStyle);
			
			HSSFCell Delegation_Exception_Rule_Flag = row.createCell(151);
			Delegation_Exception_Rule_Flag.setCellValue("Delegation Exception Rule Flag");
			Delegation_Exception_Rule_Flag.setCellStyle(headStyle);
			
			HSSFCell CTO_Line_Number = row.createCell(152);
			CTO_Line_Number.setCellValue("CTO Line Number");
			CTO_Line_Number.setCellStyle(headStyle);
			
			HSSFCell ESS_Split_Flag = row.createCell(153);
			ESS_Split_Flag.setCellValue("ESS Split Flag");
			ESS_Split_Flag.setCellStyle(headStyle);
			
			HSSFCell ESS_CTO_Line_Number = row.createCell(154);
			ESS_CTO_Line_Number.setCellValue("ESS CTO Line Number");
			ESS_CTO_Line_Number.setCellStyle(headStyle);
			
			HSSFCell CS155 = row.createCell(155);
			CS155.setCellValue("CS");
			CS155.setCellStyle(headStyle);
			
		//再来三个
			HSSFCell SP_Over_CP = row.createCell(156);
			SP_Over_CP.setCellValue("SP Over CP");
			SP_Over_CP.setCellStyle(headStyle);
			
			HSSFCell CS_SPOCP = row.createCell(157);
			CS_SPOCP.setCellValue("CS SPOCP");
			CS_SPOCP.setCellStyle(headStyle);
			
			HSSFCell ESS_SPCOP = row.createCell(158);
			ESS_SPCOP.setCellValue("ESS SPCOP");
			ESS_SPCOP.setCellStyle(headStyle);
			
		//M1------M12
			HSSFCell M1 = row.createCell(159);
			M1.setCellValue("M1");
			M1.setCellStyle(headStyle);
			
			HSSFCell M2 = row.createCell(160);
			M2.setCellValue("M2");
			M2.setCellStyle(headStyle);
			
			HSSFCell M3 = row.createCell(161);
			M3.setCellValue("M3");
			M3.setCellStyle(headStyle);
			
			HSSFCell M4 = row.createCell(162);
			M4.setCellValue("M4");
			M4.setCellStyle(headStyle);
			
			HSSFCell M5 = row.createCell(163);
			M5.setCellValue("M5");
			M5.setCellStyle(headStyle);
			
			HSSFCell M6 = row.createCell(164);
			M6.setCellValue("M6");
			M6.setCellStyle(headStyle);
			
			HSSFCell M7 = row.createCell(165);
			M7.setCellValue("M7");
			M7.setCellStyle(headStyle);
			
			HSSFCell M8 = row.createCell(166);
			M8.setCellValue("M8");
			M8.setCellStyle(headStyle);
			
			HSSFCell M9 = row.createCell(167);
			M9.setCellValue("M9");
			M9.setCellStyle(headStyle);
			
			HSSFCell M10 = row.createCell(168);
			M10.setCellValue("M10");
			M10.setCellStyle(headStyle);
			
			HSSFCell M11 = row.createCell(169);
			M11.setCellValue("M11");
			M11.setCellStyle(headStyle);
			
			HSSFCell M12 = row.createCell(170);
			M12.setCellValue("M12");
			M12.setCellStyle(headStyle);
			
	//*******************	CQ1-----CQ4
			HSSFCell CQ1 = row.createCell(171);
			CQ1.setCellValue("CQ1");
			CQ1.setCellStyle(headStyle);
			
			HSSFCell CQ2 = row.createCell(172);
			CQ2.setCellValue("CQ2");
			CQ2.setCellStyle(headStyle);
			
			HSSFCell CQ3 = row.createCell(173);
			CQ3.setCellValue("CQ3");
			CQ3.setCellStyle(headStyle);
			
			HSSFCell CQ4 = row.createCell(174);
			CQ4.setCellValue("CQ4");
			CQ4.setCellStyle(headStyle);
			
	//第一行为空的但是有样式
			HSSFCell Net_Unit_Price_null = row.createCell(175);
			Net_Unit_Price_null.setCellValue("");
			Net_Unit_Price_null.setCellStyle(headStyle);
			
			HSSFCell Net_Price_null = row.createCell(176);
			Net_Price_null.setCellValue("");
			Net_Price_null.setCellStyle(headStyle);
			
			HSSFCell Original_GTN_null = row.createCell(177);
			Original_GTN_null.setCellValue("");
			Original_GTN_null.setCellStyle(headStyle);
			
			HSSFCell New_GTN_null = row.createCell(178);
			New_GTN_null.setCellValue("");
			New_GTN_null.setCellStyle(headStyle);
			
	//最后2个
			HSSFCell PCG_Guidance = row.createCell(180);
			PCG_Guidance.setCellValue("PCG Guidance");
			PCG_Guidance.setCellStyle(headStyle);
			
			HSSFCell DCG_Guidance = row.createCell(181);
			DCG_Guidance.setCellValue("DCG Guidance");
			DCG_Guidance.setCellStyle(headStyle);
		
	/*************************************第一行标头*****************************************************/ 
	
    /*************************************第二行标头*****************************************************/ 
			HSSFCell Price = row1.createCell(10);
			Price.setCellValue("Price");
			Price.setCellStyle(headStyle);
			
			HSSFCell Margin11 = row1.createCell(11);
			Margin11.setCellValue("Margin");
			Margin11.setCellStyle(headStyle);
			
			HSSFCell PRICE12 = row1.createCell(12);
			PRICE12.setCellValue("PRICE");
			PRICE12.setCellStyle(headStyle);
			
			HSSFCell CQ_Margin13 = row1.createCell(13);
			CQ_Margin13.setCellValue("CQ Margin %");
			CQ_Margin13.setCellStyle(headStyle);
			
			HSSFCell Weighted_Margin13 = row1.createCell(14);
			Weighted_Margin13.setCellValue("Weighted Margin %");
			Weighted_Margin13.setCellStyle(headStyle);
			
			HSSFCell Quantity_CQ = row1.createCell(16);
			Quantity_CQ.setCellValue("CQ");
			Quantity_CQ.setCellStyle(headStyle);
			
			HSSFCell Quantity_CQ1 = row1.createCell(17);
			Quantity_CQ1.setCellValue("CQ+1");
			Quantity_CQ1.setCellStyle(headStyle);
			
			HSSFCell Quantity_CQ2 = row1.createCell(18);
			Quantity_CQ2.setCellValue("CQ+2");
			Quantity_CQ2.setCellStyle(headStyle);
			
			HSSFCell Quantity_CQ3 = row1.createCell(19);
			Quantity_CQ3.setCellValue("CQ+3");
			Quantity_CQ3.setCellStyle(headStyle);
			
			HSSFCell Warranty_Period = row1.createCell(29);
			Warranty_Period.setCellValue("Period");
			Warranty_Period.setCellStyle(headStyle);
			
			HSSFCell Warranty_Warranty_Type = row1.createCell(30);
			Warranty_Warranty_Type.setCellValue("Warranty Type");
			Warranty_Warranty_Type.setCellStyle(headStyle);
			
			HSSFCell Commissionable_Cost_CQ = row1.createCell(33);
			Commissionable_Cost_CQ.setCellValue("CQ");
			Commissionable_Cost_CQ.setCellStyle(headStyle);
			
			HSSFCell Commissionable_Cost_CQ1 = row1.createCell(34);
			Commissionable_Cost_CQ1.setCellValue("CQ+1");
			Commissionable_Cost_CQ1.setCellStyle(headStyle);
			
			HSSFCell Commissionable_Cost_CQ2 = row1.createCell(35);
			Commissionable_Cost_CQ2.setCellValue("CQ+2");
			Commissionable_Cost_CQ2.setCellStyle(headStyle);
			
			HSSFCell Commissionable_Cost_CQ3 = row1.createCell(36);
			Commissionable_Cost_CQ3.setCellValue("CQ+3");
			Commissionable_Cost_CQ3.setCellStyle(headStyle);
			
			//**************Cost_Adjustment
			HSSFCell Cost_Adjustment_CQ = row1.createCell(38);
			Cost_Adjustment_CQ.setCellValue("CQ");
			Cost_Adjustment_CQ.setCellStyle(headStyle);
			
			HSSFCell Cost_Adjustment_CQ1 = row1.createCell(39);
			Cost_Adjustment_CQ1.setCellValue("CQ+1");
			Cost_Adjustment_CQ1.setCellStyle(headStyle);
			
			HSSFCell Cost_Adjustment_CQ2 = row1.createCell(40);
			Cost_Adjustment_CQ2.setCellValue("CQ+2");
			Cost_Adjustment_CQ2.setCellStyle(headStyle);
			
			HSSFCell Cost_Adjustment_CQ3 = row1.createCell(41);
			Cost_Adjustment_CQ3.setCellValue("CQ+3");
			Cost_Adjustment_CQ3.setCellStyle(headStyle);
			
			HSSFCell Cost_Adjustment_Cost_Adjustment_Note = row1.createCell(42);
			Cost_Adjustment_Cost_Adjustment_Note.setCellValue("Cost Adjustment Note");
			Cost_Adjustment_Cost_Adjustment_Note.setCellStyle(headStyle);
			
			
			//Adjusted_Comm_Cost
			HSSFCell Adjusted_Comm_Cost_CQ = row1.createCell(43);
			Adjusted_Comm_Cost_CQ.setCellValue("CQ");
			Adjusted_Comm_Cost_CQ.setCellStyle(headStyle);
			
			HSSFCell Adjusted_Comm_Cost_CQ1 = row1.createCell(44);
			Adjusted_Comm_Cost_CQ1.setCellValue("CQ+1");
			Adjusted_Comm_Cost_CQ1.setCellStyle(headStyle);
			
			HSSFCell Adjusted_Comm_Cost_CQ2 = row1.createCell(45);
			Adjusted_Comm_Cost_CQ2.setCellValue("CQ+2");
			Adjusted_Comm_Cost_CQ2.setCellStyle(headStyle);
			
			HSSFCell Adjusted_Comm_Cost_CQ3 = row1.createCell(46);
			Adjusted_Comm_Cost_CQ3.setCellValue("CQ+3");
			Adjusted_Comm_Cost_CQ3.setCellStyle(headStyle);
			
			//Margin %				
			HSSFCell Margin_CQ = row1.createCell(67);
			Margin_CQ.setCellValue("CQ");
			Margin_CQ.setCellStyle(headStyle);
			
			HSSFCell Margin_CQ1 = row1.createCell(68);
			Margin_CQ1.setCellValue("CQ+1");
			Margin_CQ1.setCellStyle(headStyle);
			
			HSSFCell Margin_CQ2 = row1.createCell(69);
			Margin_CQ2.setCellValue("CQ+2");
			Margin_CQ2.setCellStyle(headStyle);
			
			HSSFCell Margin_CQ3 = row1.createCell(70);
			Margin_CQ3.setCellValue("CQ+3");
			Margin_CQ3.setCellStyle(headStyle);
			
			HSSFCell Margin_Total = row1.createCell(71);
			Margin_Total.setCellValue("Total");
			Margin_Total.setCellStyle(headStyle);
			
			//Margin $
			HSSFCell Margin_S_CQ = row1.createCell(72);
			Margin_S_CQ.setCellValue("CQ");
			Margin_S_CQ.setCellStyle(headStyle);
			
			HSSFCell Margin_S_CQ1 = row1.createCell(73);
			Margin_S_CQ1.setCellValue("CQ+1");
			Margin_S_CQ1.setCellStyle(headStyle);
			
			HSSFCell Margin_S_CQ2 = row1.createCell(74);
			Margin_S_CQ2.setCellValue("CQ+2");
			Margin_S_CQ2.setCellStyle(headStyle);
			
			HSSFCell Margin_S_CQ3 = row1.createCell(75);
			Margin_S_CQ3.setCellValue("CQ+3");
			Margin_S_CQ3.setCellStyle(headStyle);
			
			
		//TOTAL GROSS REVENUE
			HSSFCell TOTAL_GROSS_REVENUE_CQ = row1.createCell(76);
			TOTAL_GROSS_REVENUE_CQ.setCellValue("CQ");
			TOTAL_GROSS_REVENUE_CQ.setCellStyle(headStyle);
			
			HSSFCell TOTAL_GROSS_REVENUE_CQ1 = row1.createCell(77);
			TOTAL_GROSS_REVENUE_CQ1.setCellValue("CQ+1");
			TOTAL_GROSS_REVENUE_CQ1.setCellStyle(headStyle);
			
			HSSFCell TOTAL_GROSS_REVENUE_CQ2 = row1.createCell(78);
			TOTAL_GROSS_REVENUE_CQ2.setCellValue("CQ+2");
			TOTAL_GROSS_REVENUE_CQ2.setCellStyle(headStyle);
			
			HSSFCell TOTAL_GROSS_REVENUE_CQ3 = row1.createCell(79);
			TOTAL_GROSS_REVENUE_CQ3.setCellValue("CQ+3");
			TOTAL_GROSS_REVENUE_CQ3.setCellStyle(headStyle);
			
		//TOTAL_NET_REVENUE
			HSSFCell TOTAL_NET_REVENUE_CQ = row1.createCell(80);
			TOTAL_NET_REVENUE_CQ.setCellValue("CQ");
			TOTAL_NET_REVENUE_CQ.setCellStyle(headStyle);
			
			HSSFCell TOTAL_NET_REVENUE_CQ1 = row1.createCell(81);
			TOTAL_NET_REVENUE_CQ1.setCellValue("CQ+1");
			TOTAL_NET_REVENUE_CQ1.setCellStyle(headStyle);
			
			HSSFCell TOTAL_NET_REVENUE_CQ2 = row1.createCell(82);
			TOTAL_NET_REVENUE_CQ2.setCellValue("CQ+2");
			TOTAL_NET_REVENUE_CQ2.setCellStyle(headStyle);
			
			HSSFCell TOTAL_NET_REVENUE_CQ3 = row1.createCell(83);
			TOTAL_NET_REVENUE_CQ3.setCellValue("CQ+3");
			TOTAL_NET_REVENUE_CQ3.setCellStyle(headStyle);
			
		//TOTAL_Comm_COST_Inc_adjustments
			HSSFCell TOTAL_Comm_COST_Inc_adjustments_CQ = row1.createCell(84);
			TOTAL_Comm_COST_Inc_adjustments_CQ.setCellValue("CQ");
			TOTAL_Comm_COST_Inc_adjustments_CQ.setCellStyle(headStyle);
			
			HSSFCell TOTAL_Comm_COST_Inc_adjustments_CQ1 = row1.createCell(85);
			TOTAL_Comm_COST_Inc_adjustments_CQ1.setCellValue("CQ+1");
			TOTAL_Comm_COST_Inc_adjustments_CQ1.setCellStyle(headStyle);
			
			HSSFCell TOTAL_Comm_COST_Inc_adjustments_CQ2 = row1.createCell(86);
			TOTAL_Comm_COST_Inc_adjustments_CQ2.setCellValue("CQ+2");
			TOTAL_Comm_COST_Inc_adjustments_CQ2.setCellStyle(headStyle);
			
			HSSFCell TOTAL_Comm_COST_Inc_adjustmentss_CQ3 = row1.createCell(87);
			TOTAL_Comm_COST_Inc_adjustmentss_CQ3.setCellValue("CQ+3");
			TOTAL_Comm_COST_Inc_adjustmentss_CQ3.setCellStyle(headStyle);
			
		//TOTAL_MARGIN_after_revenue_reduction
			HSSFCell TOTAL_MARGIN_after_revenue_reduction_CQ = row1.createCell(88);
			TOTAL_MARGIN_after_revenue_reduction_CQ.setCellValue("CQ");
			TOTAL_MARGIN_after_revenue_reduction_CQ.setCellStyle(headStyle);
			
			HSSFCell TOTAL_MARGIN_after_revenue_reduction_CQ1 = row1.createCell(89);
			TOTAL_MARGIN_after_revenue_reduction_CQ1.setCellValue("CQ+1");
			TOTAL_MARGIN_after_revenue_reduction_CQ1.setCellStyle(headStyle);
			
			HSSFCell TOTAL_MARGIN_after_revenue_reduction_CQ2 = row1.createCell(90);
			TOTAL_MARGIN_after_revenue_reduction_CQ2.setCellValue("CQ+2");
			TOTAL_MARGIN_after_revenue_reduction_CQ2.setCellStyle(headStyle);
			
			HSSFCell TOTAL_MARGIN_after_revenue_reduction_CQ3 = row1.createCell(91);
			TOTAL_MARGIN_after_revenue_reduction_CQ3.setCellValue("CQ+3");
			TOTAL_MARGIN_after_revenue_reduction_CQ3.setCellStyle(headStyle);
		
		//Margin_after_rev_reduction
			HSSFCell Margin_after_rev_reduction_CQ = row1.createCell(92);
			Margin_after_rev_reduction_CQ.setCellValue("CQ");
			Margin_after_rev_reduction_CQ.setCellStyle(headStyle);
			
			HSSFCell Margin_after_rev_reduction_CQ1 = row1.createCell(93);
			Margin_after_rev_reduction_CQ1.setCellValue("CQ+1");
			Margin_after_rev_reduction_CQ1.setCellStyle(headStyle);
			
			HSSFCell Margin_after_rev_reduction_CQ2 = row1.createCell(94);
			Margin_after_rev_reduction_CQ2.setCellValue("CQ+2");
			Margin_after_rev_reduction_CQ2.setCellStyle(headStyle);
			
			HSSFCell Margin_after_rev_reduction_CQ3 = row1.createCell(95);
			Margin_after_rev_reduction_CQ3.setCellValue("CQ+3");
			Margin_after_rev_reduction_CQ3.setCellStyle(headStyle);
			
		//Requested
			HSSFCell Requested_Discount = row1.createCell(115);
			Requested_Discount.setCellValue("Discount");
			Requested_Discount.setCellStyle(headStyle);
			
			HSSFCell Requested_Price = row1.createCell(116);
			Requested_Price.setCellValue("Price");
			Requested_Price.setCellStyle(headStyle);
			
			HSSFCell Requested_Margin = row1.createCell(117);
			Requested_Margin.setCellValue("Margin %");
			Requested_Margin.setCellStyle(headStyle);
			
		//Entitled
			HSSFCell Entitled_Discount_From_List = row1.createCell(118);
			Entitled_Discount_From_List.setCellValue("Discount From List");
			Entitled_Discount_From_List.setCellStyle(headStyle);
			
			HSSFCell Entitled_Price = row1.createCell(119);
			Entitled_Price.setCellValue("Price");
			Entitled_Price.setCellStyle(headStyle);
			
			HSSFCell Entitled_CQ_Margin = row1.createCell(120);
			Entitled_CQ_Margin.setCellValue("Margin %");
			Entitled_CQ_Margin.setCellStyle(headStyle);
			
			HSSFCell Entitled_Weighted_Margin = row1.createCell(121);
			Entitled_Weighted_Margin.setCellValue("Weighted Margin %");
			Entitled_Weighted_Margin.setCellStyle(headStyle);
		
		//PAB_Price_
			HSSFCell PAB_Price_Discount = row1.createCell(122);
			PAB_Price_Discount.setCellValue("Discount");
			PAB_Price_Discount.setCellStyle(headStyle);
			
			HSSFCell PAB_Price_Price = row1.createCell(123);
			PAB_Price_Price.setCellValue("Price");
			PAB_Price_Price.setCellStyle(headStyle);
			
			HSSFCell PAB_Price_Margin = row1.createCell(124);
			PAB_Price_Margin.setCellValue("Margin %");
			PAB_Price_Margin.setCellStyle(headStyle);
			
		//Active_Price
			HSSFCell Active_Price_Discount = row1.createCell(125);
			Active_Price_Discount.setCellValue("Discount");
			Active_Price_Discount.setCellStyle(headStyle);
			
			HSSFCell Active_Price_Price = row1.createCell(126);
			Active_Price_Price.setCellValue("Price");
			Active_Price_Price.setCellStyle(headStyle);
			
			HSSFCell Active_Price_Margin = row1.createCell(127);
			Active_Price_Margin.setCellValue("Margin %");
			Active_Price_Margin.setCellStyle(headStyle);
		
		//Guidance_for_PWT_								
			HSSFCell Guidance_for_PWT_Guidance_for_PWT = row1.createCell(135);
			Guidance_for_PWT_Guidance_for_PWT.setCellValue("Guidance for PWT");
			Guidance_for_PWT_Guidance_for_PWT.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Discount_Band_1 = row1.createCell(136);
			Guidance_for_PWT_Discount_Band_1.setCellValue("Discount Band 1");
			Guidance_for_PWT_Discount_Band_1.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Discount_Band_2 = row1.createCell(137);
			Guidance_for_PWT_Discount_Band_2.setCellValue("Discount Band 2");
			Guidance_for_PWT_Discount_Band_2.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Discount_Band_3 = row1.createCell(138);
			Guidance_for_PWT_Discount_Band_3.setCellValue("Discount Band 3");
			Guidance_for_PWT_Discount_Band_3.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Margin_Band_1 = row1.createCell(139);
			Guidance_for_PWT_Margin_Band_1.setCellValue("Margin Band 1");
			Guidance_for_PWT_Margin_Band_1.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Margin_Band_2 = row1.createCell(140);
			Guidance_for_PWT_Margin_Band_2.setCellValue("Margin Band 2");
			Guidance_for_PWT_Margin_Band_2.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Margin_Band_3 = row1.createCell(141);
			Guidance_for_PWT_Margin_Band_3.setCellValue("Margin Band 3");
			Guidance_for_PWT_Margin_Band_3.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Discount_Color = row1.createCell(142);
			Guidance_for_PWT_Discount_Color.setCellValue("Discount Color");
			Guidance_for_PWT_Discount_Color.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Margin_Color = row1.createCell(143);
			Guidance_for_PWT_Margin_Color.setCellValue("Margin Color");
			Guidance_for_PWT_Margin_Color.setCellStyle(headStyle);
			
			HSSFCell Guidance_for_PWT_Margin_Guidance_Color = row1.createCell(144);
			Guidance_for_PWT_Margin_Guidance_Color.setCellValue("Guidance Color");
			Guidance_for_PWT_Margin_Guidance_Color.setCellStyle(headStyle);
			
		//最后几个
			HSSFCell Net_Unit_Price = row1.createCell(175);
			Net_Unit_Price.setCellValue("Guidance Color");
			Net_Unit_Price.setCellStyle(headStyle);
			
			HSSFCell Net_Price = row1.createCell(176);
			Net_Price.setCellValue("Guidance Color");
			Net_Price.setCellStyle(headStyle);
			
			HSSFCell Original_GTN = row1.createCell(177);
			Original_GTN.setCellValue("Guidance Color");
			Original_GTN.setCellStyle(headStyle);
			
			HSSFCell New_GTN = row1.createCell(178);
			New_GTN.setCellValue("Guidance Color");
			New_GTN.setCellStyle(headStyle);
			
			HSSFCell FX = row1.createCell(179);
			FX.setCellValue("");
			FX.setCellStyle(headStyle);
			
			HSSFCell FX0 = row.createCell(179);
			FX0.setCellValue("");
			FX0.setCellStyle(headStyle);
			
			HSSFCell EC_null = row.createCell(132);
			EC_null.setCellValue("");
			EC_null.setCellStyle(headStyle);
			
			HSSFCell ED_null = row.createCell(133);
			ED_null.setCellValue("");
			ED_null.setCellStyle(headStyle);
			
			HSSFCell EE_null1 = row1.createCell(134);
			EE_null1.setCellValue("");
			EE_null1.setCellStyle(headStyle);
			
			HSSFCell EC_null1 = row1.createCell(132);
			EC_null1.setCellValue("");
			EC_null1.setCellStyle(headStyle);
			
			HSSFCell ED_null1 = row1.createCell(133);
			ED_null1.setCellValue("");
			ED_null1.setCellStyle(headStyle);
			
			HSSFCell EE_null = row.createCell(134);
			EE_null.setCellValue("");
			EE_null.setCellStyle(headStyle);

	/*************************************第二行标头*****************************************************/
			
		/***************************合并单元格*************************************/
		/***************************合并单元格*************************************/				
			CellRangeAddress A = new CellRangeAddress(0, 1, 0, 0);
			sheet_1.addMergedRegion(A);
			
			CellRangeAddress B = new CellRangeAddress(0, 1, 1, 1);
			sheet_1.addMergedRegion(B);

			CellRangeAddress C = new CellRangeAddress(0, 1, 2, 2);
			sheet_1.addMergedRegion(C);

			CellRangeAddress D = new CellRangeAddress(0, 1, 3, 3);
			sheet_1.addMergedRegion(D);

			CellRangeAddress E = new CellRangeAddress(0, 1, 4, 4);
			sheet_1.addMergedRegion(E);

			CellRangeAddress F = new CellRangeAddress(0, 1, 5, 5);
			sheet_1.addMergedRegion(F);

			CellRangeAddress G = new CellRangeAddress(0, 1, 6, 6);
			sheet_1.addMergedRegion(G);

			CellRangeAddress H = new CellRangeAddress(0, 1, 7, 7);
			sheet_1.addMergedRegion(H);

			CellRangeAddress I = new CellRangeAddress(0, 1, 8, 8);
			sheet_1.addMergedRegion(I);

			CellRangeAddress J = new CellRangeAddress(0, 1, 9, 9);
			sheet_1.addMergedRegion(J);

			CellRangeAddress KL = new CellRangeAddress(0, 0, 10, 11);
			sheet_1.addMergedRegion(KL);

			CellRangeAddress MNO = new CellRangeAddress(0, 0, 12, 14);
			sheet_1.addMergedRegion(MNO);

			CellRangeAddress P = new CellRangeAddress(0, 1, 15, 15);
			sheet_1.addMergedRegion(P);

			CellRangeAddress QRST = new CellRangeAddress(0, 0, 16, 19);
			sheet_1.addMergedRegion(QRST);

			CellRangeAddress U = new CellRangeAddress(0, 1, 20, 20);
			sheet_1.addMergedRegion(U);

			CellRangeAddress V = new CellRangeAddress(0, 1, 21, 21);
			sheet_1.addMergedRegion(V);

			CellRangeAddress W = new CellRangeAddress(0, 1, 22, 22);
			sheet_1.addMergedRegion(W);

			CellRangeAddress X = new CellRangeAddress(0, 1, 23, 23);
			sheet_1.addMergedRegion(X);

			CellRangeAddress Y = new CellRangeAddress(0, 1, 24, 24);
			sheet_1.addMergedRegion(Y);

			CellRangeAddress Z = new CellRangeAddress(0, 1, 25, 25);
			sheet_1.addMergedRegion(Z);

			CellRangeAddress AA = new CellRangeAddress(0, 1, 26, 26);
			sheet_1.addMergedRegion(AA);

			CellRangeAddress AB = new CellRangeAddress(0, 1, 27, 27);
			sheet_1.addMergedRegion(AB);

			CellRangeAddress AC = new CellRangeAddress(0, 1, 28, 28);
			sheet_1.addMergedRegion(AC);

			CellRangeAddress AD_AE = new CellRangeAddress(0, 0, 29, 30);
			sheet_1.addMergedRegion(AD_AE);

			CellRangeAddress AF = new CellRangeAddress(0, 1, 31, 31);
			sheet_1.addMergedRegion(AF);

			CellRangeAddress AG = new CellRangeAddress(0, 1, 32, 32);
			sheet_1.addMergedRegion(AG);

			CellRangeAddress AH_AI_AJ_AK = new CellRangeAddress(0, 0, 33, 36);
			sheet_1.addMergedRegion(AH_AI_AJ_AK);

			CellRangeAddress AL = new CellRangeAddress(0, 1, 37, 37);
			sheet_1.addMergedRegion(AL);

			CellRangeAddress AM_AN_AO_AP_AQ = new CellRangeAddress(0, 0, 38, 42);
			sheet_1.addMergedRegion(AM_AN_AO_AP_AQ);

			CellRangeAddress AR_AS_AT_AU = new CellRangeAddress(0, 0, 43, 46);
			sheet_1.addMergedRegion(AR_AS_AT_AU);

			CellRangeAddress AV = new CellRangeAddress(0, 1, 47, 47);
			sheet_1.addMergedRegion(AV);

			CellRangeAddress AW = new CellRangeAddress(0, 1, 48, 48);
			sheet_1.addMergedRegion(AW);

			CellRangeAddress AX = new CellRangeAddress(0, 1, 49, 49);
			sheet_1.addMergedRegion(AX);

			CellRangeAddress AY = new CellRangeAddress(0, 1, 50, 50);
			sheet_1.addMergedRegion(AY);

			CellRangeAddress AZ = new CellRangeAddress(0, 1, 51, 51);
			sheet_1.addMergedRegion(AZ);

			// *******************************BBBBBB***********************************************
			CellRangeAddress BA = new CellRangeAddress(0, 1, 52, 52);
			sheet_1.addMergedRegion(BA);

			CellRangeAddress BB = new CellRangeAddress(0, 1, 53, 53);
			sheet_1.addMergedRegion(BB);

			CellRangeAddress BC = new CellRangeAddress(0, 1, 54, 54);
			sheet_1.addMergedRegion(BC);

			CellRangeAddress BD = new CellRangeAddress(0, 1, 55, 55);
			sheet_1.addMergedRegion(BD);

			CellRangeAddress BE = new CellRangeAddress(0, 1, 56, 56);
			sheet_1.addMergedRegion(BE);

			CellRangeAddress BF = new CellRangeAddress(0, 1, 57, 57);
			sheet_1.addMergedRegion(BF);

			CellRangeAddress BG = new CellRangeAddress(0, 1, 58, 58);
			sheet_1.addMergedRegion(BG);

			CellRangeAddress BH = new CellRangeAddress(0, 1, 59, 59);
			sheet_1.addMergedRegion(BH);

			CellRangeAddress BI = new CellRangeAddress(0, 1, 60, 60);
			sheet_1.addMergedRegion(BI);

			CellRangeAddress BJ = new CellRangeAddress(0, 1, 61, 61);
			sheet_1.addMergedRegion(BJ);

			CellRangeAddress BK = new CellRangeAddress(0, 1, 62, 62);
			sheet_1.addMergedRegion(BK);

			CellRangeAddress BL = new CellRangeAddress(0, 1, 63, 63);
			sheet_1.addMergedRegion(BL);

			CellRangeAddress BM = new CellRangeAddress(0, 1, 64, 64);
			sheet_1.addMergedRegion(BM);

			CellRangeAddress BN = new CellRangeAddress(0, 1, 65, 65);
			sheet_1.addMergedRegion(BN);

			CellRangeAddress BO = new CellRangeAddress(0, 1, 66, 66);
			sheet_1.addMergedRegion(BO);

			CellRangeAddress BP_BQ_BR_BS_BT = new CellRangeAddress(0, 0, 67, 71);
			sheet_1.addMergedRegion(BP_BQ_BR_BS_BT);

			CellRangeAddress BU_BV_BW_BX = new CellRangeAddress(0, 0, 72, 75);
			sheet_1.addMergedRegion(BU_BV_BW_BX);

			CellRangeAddress BY_BZ_CA_CB = new CellRangeAddress(0, 0, 76, 79);
			sheet_1.addMergedRegion(BY_BZ_CA_CB);

			CellRangeAddress CC_CD_CE_CF = new CellRangeAddress(0, 0, 80, 83);
			sheet_1.addMergedRegion(CC_CD_CE_CF);

			CellRangeAddress CG_CH_CI_CJ = new CellRangeAddress(0, 0, 84, 87);
			sheet_1.addMergedRegion(CG_CH_CI_CJ);

			CellRangeAddress CK_CL_CM_CN = new CellRangeAddress(0, 0, 88, 91);
			sheet_1.addMergedRegion(CK_CL_CM_CN);

			CellRangeAddress CO_CP_CQ_CR = new CellRangeAddress(0, 0, 92, 95);
			sheet_1.addMergedRegion(CO_CP_CQ_CR);

			CellRangeAddress CS = new CellRangeAddress(0, 1, 96, 96);
			sheet_1.addMergedRegion(CS);

			CellRangeAddress CT = new CellRangeAddress(0, 1, 97, 97);
			sheet_1.addMergedRegion(CT);

			CellRangeAddress CU = new CellRangeAddress(0, 1, 98, 98);
			sheet_1.addMergedRegion(CU);

			CellRangeAddress CV = new CellRangeAddress(0, 1, 99, 99);
			sheet_1.addMergedRegion(CV);

			CellRangeAddress CW = new CellRangeAddress(0, 1, 100, 100);
			sheet_1.addMergedRegion(CW);

			CellRangeAddress CX = new CellRangeAddress(0, 1, 101, 101);
			sheet_1.addMergedRegion(CX);

			CellRangeAddress CY = new CellRangeAddress(0, 1, 102, 102);
			sheet_1.addMergedRegion(CY);

			CellRangeAddress CZ = new CellRangeAddress(0, 1, 103, 103);
			sheet_1.addMergedRegion(CZ);

			// DA 104
			// DB 105
			// DC 106
			// DD 107
			// DE 108
			// DF 109
			// DG 110

			CellRangeAddress DH = new CellRangeAddress(0, 1, 111, 111);
			sheet_1.addMergedRegion(DH);

			CellRangeAddress DI = new CellRangeAddress(0, 1, 112, 112);
			sheet_1.addMergedRegion(DI);

			CellRangeAddress DJ = new CellRangeAddress(0, 1, 113, 113);
			sheet_1.addMergedRegion(DJ);

			CellRangeAddress DK = new CellRangeAddress(0, 1, 114, 114);
			sheet_1.addMergedRegion(DK);

			CellRangeAddress DL_DM_DN = new CellRangeAddress(0, 0, 115, 117);
			sheet_1.addMergedRegion(DL_DM_DN);

			CellRangeAddress DO_DP_DQ_DR = new CellRangeAddress(0, 0, 118, 121);
			sheet_1.addMergedRegion(DO_DP_DQ_DR);

			CellRangeAddress DS_DT_DU = new CellRangeAddress(0, 0, 122, 124);
			sheet_1.addMergedRegion(DS_DT_DU);

			CellRangeAddress DV_DW_DX = new CellRangeAddress(0, 0, 125, 127);
			sheet_1.addMergedRegion(DV_DW_DX);

			CellRangeAddress DY = new CellRangeAddress(0, 1, 128, 128);
			sheet_1.addMergedRegion(DY);

			CellRangeAddress DZ = new CellRangeAddress(0, 1, 129, 129);
			sheet_1.addMergedRegion(DZ);

			CellRangeAddress EA = new CellRangeAddress(0, 1, 130, 130);
			sheet_1.addMergedRegion(EA);

			CellRangeAddress EB = new CellRangeAddress(0, 1, 131, 131);
			sheet_1.addMergedRegion(EB);
			
			
/***************************************后加河滨单元个*******************************************************/
			
			// EC 132
			// ED 133			
			// EE 134
			CellRangeAddress EF_EO = new CellRangeAddress(0, 0, 135, 144);
			sheet_1.addMergedRegion(EF_EO);
			
			CellRangeAddress EP = new CellRangeAddress(0, 1, 145, 145);
			sheet_1.addMergedRegion(EP);
			
			CellRangeAddress EQ = new CellRangeAddress(0, 1, 146, 146);
			sheet_1.addMergedRegion(EQ);
			
			CellRangeAddress ER= new CellRangeAddress(0, 1, 147, 147);
			sheet_1.addMergedRegion(ER);
			
			CellRangeAddress ES= new CellRangeAddress(0, 1, 148, 148);
			sheet_1.addMergedRegion(ES);
			
			CellRangeAddress ET = new CellRangeAddress(0, 1, 149, 149);
			sheet_1.addMergedRegion(ET);
			
		//150------------159
			CellRangeAddress EU = new CellRangeAddress(0, 1, 150, 150);
			sheet_1.addMergedRegion(EU);
			
			CellRangeAddress EV = new CellRangeAddress(0, 1, 151, 151);
			sheet_1.addMergedRegion(EV);
			
			CellRangeAddress EW = new CellRangeAddress(0, 1, 152, 152);
			sheet_1.addMergedRegion(EW);
			
			CellRangeAddress EX = new CellRangeAddress(0, 1, 153, 153);
			sheet_1.addMergedRegion(EX);
			
			CellRangeAddress EY = new CellRangeAddress(0, 1, 154, 154);
			sheet_1.addMergedRegion(EY);
			
			CellRangeAddress EZ = new CellRangeAddress(0, 1, 155, 155);
			sheet_1.addMergedRegion(EZ);
			
			CellRangeAddress FA = new CellRangeAddress(0, 1, 156, 156);
			sheet_1.addMergedRegion(FA);
			
			CellRangeAddress FB = new CellRangeAddress(0, 1, 157, 157);
			sheet_1.addMergedRegion(FB);
			
			CellRangeAddress FC = new CellRangeAddress(0, 1, 158, 158);
			sheet_1.addMergedRegion(FC);

	   //M1------------------------------------M12
			CellRangeAddress FD = new CellRangeAddress(0, 1, 159, 159);
			sheet_1.addMergedRegion(FD);
			
			CellRangeAddress FE = new CellRangeAddress(0, 1, 160, 160);
			sheet_1.addMergedRegion(FE);
			
			CellRangeAddress FF = new CellRangeAddress(0, 1, 161, 161);
			sheet_1.addMergedRegion(FF);
			
			CellRangeAddress FG = new CellRangeAddress(0, 1, 162, 162);
			sheet_1.addMergedRegion(FG);
			
			CellRangeAddress FH = new CellRangeAddress(0, 1, 163, 163);
			sheet_1.addMergedRegion(FH);
			
			CellRangeAddress FI = new CellRangeAddress(0, 1, 164, 164);
			sheet_1.addMergedRegion(FI);
			
			CellRangeAddress FJ = new CellRangeAddress(0, 1, 165, 165);
			sheet_1.addMergedRegion(FJ);
			
			CellRangeAddress FK = new CellRangeAddress(0, 1, 166, 166);
			sheet_1.addMergedRegion(FK);
			
			CellRangeAddress FL = new CellRangeAddress(0, 1, 167, 167);
			sheet_1.addMergedRegion(FL);
			
			CellRangeAddress FM = new CellRangeAddress(0, 1, 168, 168);
			sheet_1.addMergedRegion(FM);
			
			CellRangeAddress FN = new CellRangeAddress(0, 1, 169, 169);
			sheet_1.addMergedRegion(FN);
			
			CellRangeAddress FO = new CellRangeAddress(0, 1, 170, 170);
			sheet_1.addMergedRegion(FO);
		
		//CQ1------------CQ
			CellRangeAddress FP = new CellRangeAddress(0, 1, 171, 171);
			sheet_1.addMergedRegion(FP);
			
			CellRangeAddress FQ = new CellRangeAddress(0, 1, 172, 172);
			sheet_1.addMergedRegion(FQ);
			
			CellRangeAddress FR = new CellRangeAddress(0, 1, 173, 173);
			sheet_1.addMergedRegion(FR);
			
			CellRangeAddress FS = new CellRangeAddress(0, 1, 174, 174);
			sheet_1.addMergedRegion(FS);
		
	   //175 -------- 179
			//FT	175
			//FU	176
			//FV	177
			//FW	178
			//FX	179
			
	   
	  //180--181
			CellRangeAddress FY = new CellRangeAddress(0, 1, 180, 180);
			sheet_1.addMergedRegion(FY);
			
			CellRangeAddress FZ = new CellRangeAddress(0, 1, 181, 181);
			sheet_1.addMergedRegion(FZ);
	
	/***************************************创建数据测试*******************************************************/
			//这是行数
			//一共有131个字段
			int firstData = 2;
			int dataNumber = 0;
			for (dataNumber = 0; dataNumber < list.size(); dataNumber++ ,firstData++) {
				
				//获取需要插入数据的行数
				HSSFRow firstRow = sheet_1.createRow(firstData);
				//取出每次便利的对象
				LAP_Name lap_Name = list.get(dataNumber);
				//将对象中的数据放进单元格内
			//*********-----   A--Z
				firstRow.createCell(0).setCellValue(lap_Name.getA());
				firstRow.createCell(1).setCellValue(lap_Name.getB());
				firstRow.createCell(2).setCellValue(lap_Name.getC());
				firstRow.createCell(3).setCellValue(lap_Name.getD());
				firstRow.createCell(4).setCellValue(lap_Name.getE());
				firstRow.createCell(5).setCellValue(lap_Name.getF());
				firstRow.createCell(6).setCellValue(lap_Name.getG());
				firstRow.createCell(7).setCellValue(lap_Name.getH());
				firstRow.createCell(8).setCellValue(lap_Name.getI());
				firstRow.createCell(9).setCellValue(lap_Name.getJ());
				firstRow.createCell(10).setCellValue(lap_Name.getK());
				firstRow.createCell(11).setCellValue(lap_Name.getL());
				firstRow.createCell(12).setCellValue(lap_Name.getM());
				firstRow.createCell(13).setCellValue(lap_Name.getN());
				firstRow.createCell(14).setCellValue(lap_Name.getO());
				firstRow.createCell(15).setCellValue(lap_Name.getP());
				firstRow.createCell(16).setCellValue(lap_Name.getQ());
				firstRow.createCell(17).setCellValue(lap_Name.getR());
				firstRow.createCell(18).setCellValue(lap_Name.getS());
				firstRow.createCell(19).setCellValue(lap_Name.getT());
				firstRow.createCell(20).setCellValue(lap_Name.getU());
				firstRow.createCell(21).setCellValue(lap_Name.getV());
				firstRow.createCell(22).setCellValue(lap_Name.getW());
				firstRow.createCell(23).setCellValue(lap_Name.getX());
				firstRow.createCell(24).setCellValue(lap_Name.getY());
				firstRow.createCell(25).setCellValue(lap_Name.getZ());
				
				
				
			//*********-----   AA--AZ
				firstRow.createCell(26).setCellValue(lap_Name.getAA());
				firstRow.createCell(27).setCellValue(lap_Name.getAB());
				firstRow.createCell(28).setCellValue(lap_Name.getAC());
				firstRow.createCell(29).setCellValue(lap_Name.getAD());
				firstRow.createCell(30).setCellValue(lap_Name.getAE());
				firstRow.createCell(31).setCellValue(lap_Name.getAF());
				firstRow.createCell(32).setCellValue(lap_Name.getAG());
				firstRow.createCell(33).setCellValue(lap_Name.getAH());
				firstRow.createCell(34).setCellValue(lap_Name.getAI());
				firstRow.createCell(35).setCellValue(lap_Name.getAJ());
				firstRow.createCell(36).setCellValue(lap_Name.getAK());
				firstRow.createCell(37).setCellValue(lap_Name.getAL());
				firstRow.createCell(38).setCellValue(lap_Name.getAM());
				firstRow.createCell(39).setCellValue(lap_Name.getAN());
				firstRow.createCell(40).setCellValue(lap_Name.getAO());
				firstRow.createCell(41).setCellValue(lap_Name.getAP());
				firstRow.createCell(42).setCellValue(lap_Name.getAQ());
				firstRow.createCell(43).setCellValue(lap_Name.getAR());
				firstRow.createCell(44).setCellValue(lap_Name.getAS());
				firstRow.createCell(45).setCellValue(lap_Name.getAT());
				firstRow.createCell(46).setCellValue(lap_Name.getAU());
				firstRow.createCell(47).setCellValue(lap_Name.getAV());
				firstRow.createCell(48).setCellValue(lap_Name.getAW());
				firstRow.createCell(49).setCellValue(lap_Name.getAX());
				firstRow.createCell(50).setCellValue(lap_Name.getAY());
				firstRow.createCell(51).setCellValue(lap_Name.getAZ());
				
			//*********-----   BA--BZ
				firstRow.createCell(52).setCellValue(lap_Name.getBA());
				firstRow.createCell(53).setCellValue(lap_Name.getBB());
				firstRow.createCell(54).setCellValue(lap_Name.getBC());
				firstRow.createCell(55).setCellValue(lap_Name.getBD());
				firstRow.createCell(56).setCellValue(lap_Name.getBE());
				firstRow.createCell(57).setCellValue(lap_Name.getBF());
				firstRow.createCell(58).setCellValue(lap_Name.getBG());
				firstRow.createCell(59).setCellValue(lap_Name.getBH());
				firstRow.createCell(60).setCellValue(lap_Name.getBI());
				firstRow.createCell(61).setCellValue(lap_Name.getBJ());
				firstRow.createCell(62).setCellValue(lap_Name.getBK());
				firstRow.createCell(63).setCellValue(lap_Name.getBL());
				firstRow.createCell(64).setCellValue(lap_Name.getBM());
				firstRow.createCell(65).setCellValue(lap_Name.getBN());
				firstRow.createCell(66).setCellValue(lap_Name.getBO());
				firstRow.createCell(67).setCellValue(lap_Name.getBP());
				firstRow.createCell(68).setCellValue(lap_Name.getBQ());
				firstRow.createCell(69).setCellValue(lap_Name.getBR());
				firstRow.createCell(70).setCellValue(lap_Name.getBS());
				firstRow.createCell(71).setCellValue(lap_Name.getBT());
				firstRow.createCell(72).setCellValue(lap_Name.getBU());
				firstRow.createCell(73).setCellValue(lap_Name.getBV());
				firstRow.createCell(74).setCellValue(lap_Name.getBW());
				firstRow.createCell(75).setCellValue(lap_Name.getBX());
				firstRow.createCell(76).setCellValue(lap_Name.getBY());
				firstRow.createCell(77).setCellValue(lap_Name.getBZ());
				
				
		//*********-----   CA--CZ
				firstRow.createCell(78).setCellValue(lap_Name.getCA());
				firstRow.createCell(79).setCellValue(lap_Name.getCB());
				firstRow.createCell(80).setCellValue(lap_Name.getCC());
				firstRow.createCell(81).setCellValue(lap_Name.getCD());
				firstRow.createCell(82).setCellValue(lap_Name.getCE());
				firstRow.createCell(83).setCellValue(lap_Name.getCF());
				firstRow.createCell(84).setCellValue(lap_Name.getCG());
				firstRow.createCell(85).setCellValue(lap_Name.getCH());
				firstRow.createCell(86).setCellValue(lap_Name.getCI());
				firstRow.createCell(87).setCellValue(lap_Name.getCJ());
				firstRow.createCell(88).setCellValue(lap_Name.getCK());
				firstRow.createCell(89).setCellValue(lap_Name.getCL());
				firstRow.createCell(90).setCellValue(lap_Name.getCM());
				firstRow.createCell(91).setCellValue(lap_Name.getCN());
				firstRow.createCell(92).setCellValue(lap_Name.getCO());
				firstRow.createCell(93).setCellValue(lap_Name.getCP());
				firstRow.createCell(94).setCellValue(lap_Name.getCQ());
				firstRow.createCell(95).setCellValue(lap_Name.getCR());
				firstRow.createCell(96).setCellValue(lap_Name.getCS());
				firstRow.createCell(97).setCellValue(lap_Name.getCT());
				firstRow.createCell(98).setCellValue(lap_Name.getCU());
				firstRow.createCell(99).setCellValue(lap_Name.getCV());
				firstRow.createCell(100).setCellValue(lap_Name.getCW());
				firstRow.createCell(101).setCellValue(lap_Name.getCX());
				firstRow.createCell(102).setCellValue(lap_Name.getCY());
				firstRow.createCell(103).setCellValue(lap_Name.getCZ());
				firstRow.createCell(104).setCellValue(lap_Name.getCZ());
				
			//*********-----   DA--DZ
				firstRow.createCell(104).setCellValue(lap_Name.getDA());
				firstRow.createCell(105).setCellValue(lap_Name.getDB());
				firstRow.createCell(106).setCellValue(lap_Name.getDC());
				firstRow.createCell(107).setCellValue(lap_Name.getDD());
				firstRow.createCell(108).setCellValue(lap_Name.getDE());
				firstRow.createCell(109).setCellValue(lap_Name.getDF());
				firstRow.createCell(110).setCellValue(lap_Name.getDG());
				firstRow.createCell(111).setCellValue(lap_Name.getDH());
				firstRow.createCell(112).setCellValue(lap_Name.getDI());
				firstRow.createCell(113).setCellValue(lap_Name.getDJ());
				firstRow.createCell(114).setCellValue(lap_Name.getDK());
				firstRow.createCell(115).setCellValue(lap_Name.getDL());
				firstRow.createCell(116).setCellValue(lap_Name.getDM());
				firstRow.createCell(117).setCellValue(lap_Name.getDN());
				firstRow.createCell(118).setCellValue(lap_Name.getDO());
				firstRow.createCell(119).setCellValue(lap_Name.getDP());
				firstRow.createCell(120).setCellValue(lap_Name.getDQ());
				firstRow.createCell(121).setCellValue(lap_Name.getDR());
				firstRow.createCell(122).setCellValue(lap_Name.getDS());
				firstRow.createCell(123).setCellValue(lap_Name.getDT());
				firstRow.createCell(124).setCellValue(lap_Name.getDU());
				firstRow.createCell(125).setCellValue(lap_Name.getDV());
				firstRow.createCell(126).setCellValue(lap_Name.getDW());
				firstRow.createCell(127).setCellValue(lap_Name.getDX());
				firstRow.createCell(128).setCellValue(lap_Name.getDY());
				firstRow.createCell(129).setCellValue(lap_Name.getDZ());
				
			//*********-----   EA--EZ
				firstRow.createCell(130).setCellValue(lap_Name.getEA());
				firstRow.createCell(131).setCellValue(lap_Name.getEB());
				firstRow.createCell(132).setCellValue(lap_Name.getEC());
				firstRow.createCell(133).setCellValue(lap_Name.getED());
				firstRow.createCell(134).setCellValue(lap_Name.getEE());
				firstRow.createCell(135).setCellValue(lap_Name.getEF());
				firstRow.createCell(136).setCellValue(lap_Name.getEG());
				firstRow.createCell(137).setCellValue(lap_Name.getEH());
				firstRow.createCell(138).setCellValue(lap_Name.getEI());
				firstRow.createCell(139).setCellValue(lap_Name.getEJ());
				firstRow.createCell(140).setCellValue(lap_Name.getEK());
				firstRow.createCell(141).setCellValue(lap_Name.getEL());
				firstRow.createCell(142).setCellValue(lap_Name.getEM());
				firstRow.createCell(143).setCellValue(lap_Name.getEN());
				firstRow.createCell(144).setCellValue(lap_Name.getEO());
				firstRow.createCell(145).setCellValue(lap_Name.getEP());
				firstRow.createCell(146).setCellValue(lap_Name.getEQ());
				firstRow.createCell(147).setCellValue(lap_Name.getER());
				firstRow.createCell(148).setCellValue(lap_Name.getES());
				firstRow.createCell(149).setCellValue(lap_Name.getET());
				firstRow.createCell(150).setCellValue(lap_Name.getEU());
				firstRow.createCell(151).setCellValue(lap_Name.getEV());
				firstRow.createCell(152).setCellValue(lap_Name.getEW());
				firstRow.createCell(153).setCellValue(lap_Name.getEX());
				firstRow.createCell(154).setCellValue(lap_Name.getEY());
				firstRow.createCell(155).setCellValue(lap_Name.getEZ());
				
				
			//*********-----   FA--FZ
				firstRow.createCell(156).setCellValue(lap_Name.getFA());
				firstRow.createCell(157).setCellValue(lap_Name.getFB());
				firstRow.createCell(158).setCellValue(lap_Name.getFC());
				firstRow.createCell(159).setCellValue(lap_Name.getFD());
				firstRow.createCell(160).setCellValue(lap_Name.getFE());
				firstRow.createCell(161).setCellValue(lap_Name.getFF());
				firstRow.createCell(162).setCellValue(lap_Name.getFG());
				firstRow.createCell(163).setCellValue(lap_Name.getFH());
				firstRow.createCell(164).setCellValue(lap_Name.getFI());
				firstRow.createCell(165).setCellValue(lap_Name.getFJ());
				firstRow.createCell(166).setCellValue(lap_Name.getFK());
				firstRow.createCell(167).setCellValue(lap_Name.getFL());
				firstRow.createCell(168).setCellValue(lap_Name.getFM());
				firstRow.createCell(169).setCellValue(lap_Name.getFN());
				firstRow.createCell(170).setCellValue(lap_Name.getFO());
				firstRow.createCell(171).setCellValue(lap_Name.getFP());
				firstRow.createCell(172).setCellValue(lap_Name.getFQ());
				firstRow.createCell(173).setCellValue(lap_Name.getFR());
				firstRow.createCell(174).setCellValue(lap_Name.getFS());
				firstRow.createCell(175).setCellValue(lap_Name.getFT());
				firstRow.createCell(176).setCellValue(lap_Name.getFU());
				firstRow.createCell(177).setCellValue(lap_Name.getFV());
				firstRow.createCell(178).setCellValue(lap_Name.getFW());
				firstRow.createCell(179).setCellValue(lap_Name.getFX());
				firstRow.createCell(180).setCellValue(lap_Name.getFY());
				firstRow.createCell(181).setCellValue(lap_Name.getFZ());
				
			}
			System.out.println("插入了----"+dataNumber+"条------数据");
			
			return wb;
		} catch (Exception e) {
			e.getStackTrace();
		}
		return null;
	}
	/********************************************************************************/
		/**
		 * @deprecated 标头设置
		 * @param wb
		 * @return CellStyle
		 */
		private static CellStyle headStyle(HSSFWorkbook wb) {
			// 背景颜色
			CellStyle createCellStyle = wb.createCellStyle();
			
			createCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//			createCellStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			createCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			
			//设置自动换行:  
			createCellStyle.setWrapText(true);//设置自动换行 
			
			// 边框颜色
			createCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框    
			createCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框    
			createCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
			createCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框    
			createCellStyle.setTopBorderColor(HSSFColor.BLACK.index);
			createCellStyle.setBottomBorderColor(HSSFColor.BLACK.index);
			createCellStyle.setLeftBorderColor(HSSFColor.BLACK.index);
			createCellStyle.setRightBorderColor(HSSFColor.BLACK.index);
			
	
			createCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			return createCellStyle;
		}
		
	/********************************************************************************/

}
