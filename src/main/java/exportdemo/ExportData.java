package exportdemo;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportData {
	// CONSTANTS
	/*----------------------------------------------------------*/
	static final String FILE_SAVE_LOCATION = "C:\\reports\\";
	static final String FILE_NAME = "Test report.xlsx";
	/*----------------------------------------------------------*/

	public static void main(String[] args) throws IOException {

		// creating workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		// creating sheet with name "Report" in workbook
		XSSFSheet sheet = workbook.createSheet("Report");

		XSSFDrawing drawing = sheet.createDrawingPatriarch();

		XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 10, 20, 30, 40);

		XSSFChart chart = drawing.createChart(anchor);
		chart.setTitleText("Test results");
		chart.setTitleOverlay(false);

		XDDFChartLegend legend = chart.getOrAddLegend();
		legend.setPosition(LegendPosition.TOP_RIGHT);
		String[] legendData = { "passed", "skipped", "failed" };
		XDDFDataSource<String> testOutcomes = XDDFDataSourcesFactory.fromArray(legendData);
		Integer[] numericData = { 10, 12, 30 };
		XDDFNumericalDataSource<Integer> values = XDDFDataSourcesFactory.fromArray(numericData);

		XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);// for simple pie chart you can use
																			// ChartTypes.PIE
		chart.displayBlanksAs(null);
		data.setVaryColors(true);
		data.addSeries(testOutcomes, values);

		chart.plot(data);

		try (FileOutputStream outputStream = new FileOutputStream(FILE_SAVE_LOCATION + FILE_NAME)) {
			workbook.write(outputStream);
		} finally {
			// don't forget to close workbook to prevent memory leaks
			workbook.close();
		}
	}
}
