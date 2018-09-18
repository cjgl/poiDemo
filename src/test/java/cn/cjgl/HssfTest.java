package cn.cjgl;

import static org.junit.Assert.*;

import java.io.IOException;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

public class HssfTest {
	
	private Hssf hssf;

	@BeforeClass
	public static void setUpBeforeClass() throws Exception {
	}

	@AfterClass
	public static void tearDownAfterClass() throws Exception {
	}

	@Before
	public void setUp() throws Exception {
		this.hssf = new Hssf();
	}

	@After
	public void tearDown() throws Exception {
	}

	@Test
	public void testCreateCells() throws IOException {
		this.hssf.createCells();
	}

	@Test
	public void testDataValidation() throws IOException {
		this.hssf.dataValidation();
	}

	@Test
	public void testForEach() throws IOException {
		this.hssf.forEach();
	}

	@Test
	public void testGetText() throws IOException {
		this.hssf.getText();
	}

	@Test
	public void testGroupRow() throws IOException {
		this.hssf.groupRow();
	}

	@Test
	public void testMergedRegion() throws IOException {
		this.hssf.mergedRegion();
	}

	@Test
	public void testSetBackgroundColor() throws IOException {
		this.hssf.setBackgroundColor();
	}

	@Test
	public void testSetComment() throws IOException {
		this.hssf.setComment();
	}

	@Test
	public void testSetFooter() throws IOException {
		this.hssf.setFooter();
	}

	@Test
	public void testTestFont() throws IOException {
		this.hssf.testFont();
	}

	@Test
	public void testTestHyperlink() throws IOException {
		this.hssf.testHyperlink();
	}

	@Test
	public void testHssfAlign() throws IOException {
		this.hssf.hssfAlign();
	}

}
