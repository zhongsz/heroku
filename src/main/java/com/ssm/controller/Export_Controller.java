package com.ssm.controller;

import java.io.OutputStream;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.record.WSBoolRecord;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
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

/**
 * 
 * @author Administrator
 *
 */
@Controller
public class Export_Controller {

	@ResponseBody
	@RequestMapping("/LAP_Export")
	public void export(HttpServletRequest request, HttpServletResponse response) {

		/*******************************************************************************/

		Object list = null;
		HSSFWorkbook wb = Export_Controller.export_text(list);

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
	@SuppressWarnings("deprecation")
	private static HSSFWorkbook export_text(Object list) {
		try {
			// 创建工作薄对象
			HSSFWorkbook wb = new HSSFWorkbook();

			// 创建sheet页并命名
			HSSFSheet sheet_1 = wb.createSheet("Apttus_Proposal__Proposal");
			HSSFRow row = sheet_1.createRow(0);
			HSSFRow row1 = sheet_1.createRow(1);
			HSSFRow row2 = sheet_1.createRow(2);
			// 创建标题行样式设置样式
			// 背景颜色
			CellStyle createCellStyle = wb.createCellStyle();
			createCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			createCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

			// 边框颜色
			createCellStyle.setBorderBottom(IndexedColors.BLACK.getIndex());
			createCellStyle.setBorderLeft(IndexedColors.BLACK.getIndex());
			createCellStyle.setBorderRight(IndexedColors.BLACK.getIndex());
			createCellStyle.setBorderTop(IndexedColors.BLACK.getIndex());
			createCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

			createCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			
//			row.setRowStyle(createCellStyle);
//			row1.setRowStyle(createCellStyle);
//			row2.setRowStyle(createCellStyle);

			// 合并单元格
			// CellRangeAddress cra =new CellRangeAddress(1, 3, 1, 3); // 起始行, 终止行, 起始列, 终止列
			CellRangeAddress A = new CellRangeAddress(0, 1, 0, 0);
			
			sheet_1.addMergedRegion(A);
			
			HSSFRow rowA = sheet_1.getRow(0);
			HSSFCell cell = rowA.getCell(0);
			cell.setCellValue("AA");
			cell.setCellStyle(createCellStyle);

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
			
			/*****************************创建标头内容*********************************************/
			
			//第一行别称为row
			//第二行别成为row1
//			HSSFCell NP_Flag = row.getCell(0);
//			NP_Flag.setCellValue("NP Flag");
//			NP_Flag.setCellStyle(createCellStyle);

			
			
			
			/*****************************创建标头内容*********************************************/
			
			return wb;
		} catch (Exception e) {
			e.getStackTrace();
		}
		return null;
	}
	/********************************************************************************/

}
