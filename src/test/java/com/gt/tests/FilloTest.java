package com.gt.tests;

import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class FilloTest {

	Connection connection;
	Fillo fillo = new Fillo();

	@BeforeTest
	public void beforeTest() {
		try {

			connection = fillo.getConnection(System.getProperty("user.dir") + "//data//Name_Details.xlsx");

		} catch (FilloException e) {

			e.printStackTrace();
		}
	}

	@Test
	public void extractData() {

		try {

			String strQuery = "Select * from Sheet1";
			Recordset recordset = connection.executeQuery(strQuery);

			while (recordset.next()) {
				System.out.println(recordset.getField("Name"));
				System.out.println(recordset.getField("Age"));
				System.out.println(recordset.getField("Sex"));
			}

			recordset.close();
			connection.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	@Test
	public void writeData() {

		try {

			String strQuery = "Update Sheet1 Set Age=35 where ID=1 and name='Giri'";

			connection.executeUpdate(strQuery);

			connection.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void addData() {

		try {

			String strQuery = "INSERT INTO Sheet1(Name,Age,Sex) VALUES('Peter','50','M')";

			connection.executeUpdate(strQuery);

			connection.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void removeData() {

		try {

			String strQuery = "DELETE FROM Sheet1 WHERE Name='Peter'";

			connection.executeUpdate(strQuery);

			connection.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
