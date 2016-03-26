package br.net.felipeweb.excel.cep.manipulator;

import br.com.postmon.jpostmon.Consultas;
import br.com.postmon.jpostmon.Postmon;
import br.com.postmon.jpostmon.dao.Endereco;
import br.com.postmon.jpostmon.exception.PostmonAPIException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.security.InvalidParameterException;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelManipulator {
	private final File file;
	private final String sheet;
	private final String cepColl;
	private XSSFWorkbook workbook;

	public ExcelManipulator(File file, String sheet, String cepColl) {
		this.file = file;
		this.sheet = sheet;
		this.cepColl = cepColl;
	}

	public void getAddress() throws IOException, InvalidFormatException {
		getCepFromSheet();
		String name = file.getName().split(".xlsx")[0];
		String tmpFileName = name + "_CEP.xlsx";
		OutputStream fileOutputStream = new FileOutputStream(tmpFileName);
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		workbook.close();
		new File(tmpFileName).delete();
	}


	private void getCepFromSheet() throws IOException, InvalidFormatException {
		workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet(this.sheet);
		int numRows = sheet.getPhysicalNumberOfRows();
		for (int i = 1; i < numRows; i++) {
			CellReference ref = new CellReference(cepColl + i);
			XSSFRow row = sheet.getRow(i);
			XSSFCell cell = row.getCell(ref.getCol());
			String cep = cell.getRawValue();
			Endereco endereco = findCep(cep);
			short lastCellNum = row.getLastCellNum();
			if (endereco != null) {
				XSSFCell rowCellRua = row.createCell(lastCellNum);
				rowCellRua.setCellValue(endereco.getLogradouro());
				int rowCellRuaColumnIndex = rowCellRua.getColumnIndex();
				sheet.autoSizeColumn(rowCellRuaColumnIndex);
				XSSFCell bairroCell = row.createCell(lastCellNum + 1);
				bairroCell.setCellValue(endereco.getBairro());
				int bairroCellColumnIndex = bairroCell.getColumnIndex();
				sheet.autoSizeColumn(bairroCellColumnIndex);
				XSSFCell cidadeCell = row.createCell(lastCellNum + 2);
				cidadeCell.setCellValue(endereco.getCidade());
				int cidadeCellColumnIndex = cidadeCell.getColumnIndex();
				sheet.autoSizeColumn(cidadeCellColumnIndex);
				XSSFCell estadoCell = row.createCell(lastCellNum + 3);
				estadoCell.setCellValue(endereco.getEstado());
				int estadoCellColumnIndex = estadoCell.getColumnIndex();
				sheet.autoSizeColumn(estadoCellColumnIndex);
				if (sheet.getRow(0).getCell(rowCellRuaColumnIndex) == null) {
					XSSFFont font = workbook.createFont();
					font.setFontName("Arial");
					font.setBold(true);
					XSSFCellStyle cellStyle = workbook.createCellStyle();
					cellStyle.setAlignment(HorizontalAlignment.CENTER);
					cellStyle.setFont(font);
					XSSFCell titleRua = sheet.getRow(0).createCell(rowCellRuaColumnIndex);
					titleRua.setCellStyle(cellStyle);
					titleRua.setCellValue("RUA");
					XSSFCell titleBairro = sheet.getRow(0).createCell(bairroCellColumnIndex);
					titleBairro.setCellStyle(cellStyle);
					titleBairro.setCellValue("BAIRRO");
					XSSFCell cidadeTitle = sheet.getRow(0).createCell(cidadeCellColumnIndex);
					cidadeTitle.setCellStyle(cellStyle);
					cidadeTitle.setCellValue("CIDADE");
					XSSFCell estadoTitle = sheet.getRow(0).createCell(estadoCellColumnIndex);
					estadoTitle.setCellStyle(cellStyle);
					estadoTitle.setCellValue("ESTADO");
				}
			}
		}
	}

	private Endereco findCep(String cep) {
		Endereco endereco = null;
		try {
			endereco = Postmon.consultar(Consultas.CEP).cep(cep).buscar();
			System.out.println(endereco.toString());
		} catch (PostmonAPIException e) {
			System.out.println(e.getMessage());
			try {
				endereco = Postmon.consultar(Consultas.CEP).cep(cep).buscar();
				System.out.println(endereco.toString());
			} catch (PostmonAPIException e1) {
				System.out.println(e1.getMessage());
			}
		} catch (InvalidParameterException e) {
			System.out.println(e.getMessage());
		}
		return endereco;
	}
}
