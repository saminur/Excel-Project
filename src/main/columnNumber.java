package main;

public class columnNumber {

	public int stringToColumn(String a) {
		if (a.equals(""))
			return 0;
		int stringSize = a.length();
		int columnNumber = 0;

		if (stringSize > 1) {

			for (int i = 0; i < stringSize - 1; i++) {
				int val = (int) Math.pow(26, stringSize - 1 - i);
				int b = (int) a.charAt(i);
				int bb = b - 65 + 1;
				val = val * bb;
				columnNumber += val;
			}

			int aaa = (int) a.charAt(stringSize - 1);
			// System.out.println(aaa);
			int aa = aaa - 65;
			// System.out.println(aa);
			columnNumber = columnNumber + aa + 1;
			// System.out.println(columnNumber);

		} else {
			int aaa = (int) a.charAt(0);
			int aa = aaa - 65;
			columnNumber = aa + 1;
			// System.out.println(columnNumber);
		}
		return columnNumber;
	}

}
