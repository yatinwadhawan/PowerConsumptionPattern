import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Array;
import java.util.ArrayList;

import jxl.Workbook;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Main {

	public static void main(String[] str) throws IOException {

		String firstLine = "DATA (TSContingencyElement, [WhoAmI:2,TSTimeInCycles,TSTimeInSeconds,WhoAmI,"
				+ "TSEventString,Enabled,FilterName,Comment,TSCTGName]) \n{\n";
		String middle = "";
		String lastline = "\n}";

		String csvFile = "Workbook1.csv";
		BufferedReader br = null;
		String line = "";
		String cvsSplitBy = ",";

		ArrayList<PowerDemand> lp = new ArrayList<PowerDemand>();

		try {
			br = new BufferedReader(new FileReader(csvFile));
			while ((line = br.readLine()) != null) {

				String[] country = line.split(cvsSplitBy);
				PowerDemand p = new PowerDemand(country[0], country[1]);
				lp.add(p);
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		int size = lp.size();
		StringBuffer buffer = new StringBuffer();
		double time = 1.0000;
		int count = 0;
		for (int i = 0; i < size / 2; i++) {
			PowerDemand p = lp.get(i);
			double demand;
			if (count < 2000)
				demand = Double.parseDouble(p.getDemand()) / (3 * 100);
			else
				demand = -Double.parseDouble(p.getDemand()) / (3 * 100);

			count++;
			double cycle = time * 60.0000;
			time = time + 10.0000;
			String load5 = "\"Load Bus 5 #1\"    " + cycle + " " + time
					+ " \"Load '5' '1'\" \"CHANGEBY " + demand + " MW\" "
					+ "\"CHECK\" \"\" \"\" \"My Transient Contingency\"";
			String load6 = "\"Load Bus 6 #1\"   " + cycle + " " + time
					+ " \"Load '6' '1'\" \"CHANGEBY " + demand + " MW\" "
					+ "\"CHECK\" \"\" \"\" \"My Transient Contingency\"";
			String load8 = "\"Load Bus 8 #1\"    " + cycle + " " + time
					+ " \"Load '8' '1'\" \"CHANGEBY " + demand + " MW\" "
					+ "\"CHECK\" \"\" \"\" \"My Transient Contingency\"";
			buffer.append("\n" + load5 + "\n" + load6 + "\n" + load8);
		}

		File file = new File(
				"/Users/yatinwadhawan/Documents/Security/Research/Data Set/microgrid/Data/data.aux");
		file.createNewFile();
		FileWriter writer = new FileWriter(file);
		writer.write(firstLine + buffer.toString() + lastline);
		writer.flush();
		writer.close();
	}

	// Logic to collect the demand in one file.
	public static void collectDemand() throws IOException,
			RowsExceededException, WriteException {
		final File folder = new File(
				"/Users/yatinwadhawan/Documents/Security/Research/Data Set/microgrid/");
		ArrayList<String> ls = new ArrayList<String>();
		ls = listFilesForFolder(folder);

		ArrayList<ArrayList<PowerDemand>> list = new ArrayList<ArrayList<PowerDemand>>();

		for (int i = 1; i < ls.size() - 2; i++) {
			String csvFile = "/Users/yatinwadhawan/Documents/Security/Research/Data Set/microgrid/"
					+ ls.get(i);
			// System.out.println(csvFile);

			BufferedReader br = null;
			String line = "";
			String cvsSplitBy = ",";

			try {

				ArrayList<PowerDemand> lp = new ArrayList<PowerDemand>();
				br = new BufferedReader(new FileReader(csvFile));
				while ((line = br.readLine()) != null) {

					String[] country = line.split(cvsSplitBy);
					PowerDemand p = new PowerDemand(country[0], country[1]);
					lp.add(p);
				}
				list.add(lp);

			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		System.out.println(list.size());

		// Adding all the timestamps
		WritableWorkbook wworkbook = Workbook
				.createWorkbook(new File(
						"/Users/yatinwadhawan/Documents/Security/Research/Data Set/microgrid/Data/output.xls"));
		WritableSheet wsheet = wworkbook.createSheet("First Sheet", 0);
		int row = 0, column = 0;
		ArrayList<PowerDemand> p = list.get(0);
		for (int i = 0; i < p.size(); i++) {
			Number num = new Number(column, row++, Double.parseDouble(p.get(i)
					.getTimestamp()));
			wsheet.addCell(num);
		}
		column++;
		for (int j = 1; j < 250; j++) {
			row = 0;
			ArrayList<PowerDemand> val = list.get(j);
			for (int i = 0; i < val.size(); i++) {
				Number num = new Number(column, row++, Double.parseDouble(val
						.get(i).getDemand()));
				wsheet.addCell(num);
			}
			System.out.println(column);
			column++;
		}

		wworkbook.write();
		wworkbook.close();

		WritableWorkbook wworkbook1 = Workbook
				.createWorkbook(new File(
						"/Users/yatinwadhawan/Documents/Security/Research/Data Set/microgrid/Data/output2.xls"));
		WritableSheet wsheet1 = wworkbook1.createSheet("First Sheet", 0);
		int row1 = 0, column1 = 0;
		ArrayList<PowerDemand> p1 = list.get(0);
		for (int i = 0; i < p1.size(); i++) {
			Number num = new Number(column1, row1++, Double.parseDouble(p1.get(
					i).getTimestamp()));
			wsheet1.addCell(num);
		}
		column1++;

		for (int j = 250; j < 443; j++) {
			row1 = 0;
			ArrayList<PowerDemand> val = list.get(j);
			for (int i = 0; i < val.size(); i++) {
				Number num = new Number(column1, row1++, Double.parseDouble(val
						.get(i).getDemand()));
				wsheet1.addCell(num);
			}
			System.out.println(column1);
			column1++;
		}
		wworkbook1.write();
		wworkbook1.close();
	}

	// Reading name of the files from the folder.
	public static ArrayList<String> listFilesForFolder(final File folder) {
		ArrayList<String> ls = new ArrayList<String>();
		for (final File fileEntry : folder.listFiles()) {
			if (fileEntry.isDirectory()) {
				listFilesForFolder(fileEntry);
			} else {
				ls.add(fileEntry.getName());
			}
		}
		return ls;
	}
}