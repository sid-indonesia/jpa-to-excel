package org.sidindonesia.jpatoexcel.util;

public final class CamelCaseUtil {

	private CamelCaseUtil() {
		// Utility class
	}

	// Function to convert camel case
	// string to snake case string
	public static String camelToSnake(String str) {
		// Regular Expression
		String regex = "([a-z])([A-Z]+)";

		// Replacement string
		String replacement = "$1_$2";

		// Replace the given regex
		// with replacement string
		// and convert it to lower case.
		str = str.replaceAll(regex, replacement).toLowerCase();

		// return string
		return str;
	}
}
