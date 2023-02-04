package DataDriven.Excel;

import java.io.IOException;
import java.util.ArrayList;

public class Testsample {

	public static void main(String[] args) throws IOException {
		Excel ex = new Excel();
		ArrayList data = ex.getData("Purchase");
		System.out.println(data.get(0));
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));

	}

}
