<?php

/**
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2011
 * @license LGPLv3, http://www.arcavias.com/en/license
 */


class Controller_ExtJS_Product_Import_Text_ExcelTest extends MW_Unittest_Testcase
{
	private $_object;
	private $_context;


	/**
	 * Runs the test methods of this class.
	 *
	 * @access public
	 * @static
	 */
	public static function main()
	{
		require_once 'PHPUnit/TextUI/TestRunner.php';

		$suite  = new PHPUnit_Framework_TestSuite( 'Controller_ExtJS_Product_Import_Text_ExcelTest' );
		$result = PHPUnit_TextUI_TestRunner::run( $suite );
	}


	/**
	 * Sets up the fixture, for example, opens a network connection.
	 * This method is called before a test is executed.
	 *
	 * @access protected
	 */
	protected function setUp()
	{
		$this->_context = TestHelper::getContext();
		$this->_context->getConfig()->set( 'controller/extjs/product/export/text/default/container/type', 'PHPExcel' );
		$this->_context->getConfig()->set( 'controller/extjs/product/export/text/default/container/format', 'Excel5' );
		$this->_context->getConfig()->set( 'controller/extjs/product/import/text/default/container/type', 'PHPExcel' );
		$this->_context->getConfig()->set( 'controller/extjs/product/import/text/default/container/format', 'Excel5' );

		$this->_object = new Controller_ExtJS_Product_Import_Text_Default( $this->_context );
	}


	/**
	 * Tears down the fixture, for example, closes a network connection.
	 * This method is called after a test is executed.
	 *
	 * @access protected
	 */
	protected function tearDown()
	{
		$this->_object = null;

		Controller_ExtJS_Factory::clear();
		MShop_Factory::clear();
	}


	public function testImportFromXLSFile()
	{
		$this->_object = new Controller_ExtJS_Product_Import_Text_Default( $this->_context );

		$filename = 'product-import-test.xlsx';

		$phpExcel = new PHPExcel();
		$phpExcel->setActiveSheetIndex(0);
		$sheet = $phpExcel->getActiveSheet();

		$sheet->setCellValueByColumnAndRow( 0, 2, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 3, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 4, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 5, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 6, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 7, 'en' );

		$sheet->setCellValueByColumnAndRow( 1, 2, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 3, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 4, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 5, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 6, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 7, 'product' );

		$sheet->setCellValueByColumnAndRow( 2, 2, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 3, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 4, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 5, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 6, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 7, 'ABCD' );

		$sheet->setCellValueByColumnAndRow( 3, 2, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 3, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 4, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 5, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 6, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 7, 'default' );

		$sheet->setCellValueByColumnAndRow( 4, 2, 'long' );
		$sheet->setCellValueByColumnAndRow( 4, 3, 'metadescription' );
		$sheet->setCellValueByColumnAndRow( 4, 4, 'metakeywords' );
		$sheet->setCellValueByColumnAndRow( 4, 5, 'metatitle' );
		$sheet->setCellValueByColumnAndRow( 4, 6, 'name' );
		$sheet->setCellValueByColumnAndRow( 4, 7, 'short' );

		$sheet->setCellValueByColumnAndRow( 6, 2, 'ABCD: long' );
		$sheet->setCellValueByColumnAndRow( 6, 3, 'ABCD: meta desc' );
		$sheet->setCellValueByColumnAndRow( 6, 4, 'ABCD: meta keywords' );
		$sheet->setCellValueByColumnAndRow( 6, 5, 'ABCD: meta title' );
		$sheet->setCellValueByColumnAndRow( 6, 6, 'ABCD: name' );
		$sheet->setCellValueByColumnAndRow( 6, 7, 'ABCD: short' );

		$objWriter = PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');
		$objWriter->save($filename);


		$params = new stdClass();
		$params->site = $this->_context->getLocale()->getSite()->getCode();
		$params->items = $filename;

		$this->_object->importFile( $params );


		$textManager = MShop_Text_Manager_Factory::createManager( $this->_context );
		$criteria = $textManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'text.domain', 'product' );
		$expr[] = $criteria->compare( '==', 'text.languageid', 'en' );
		$expr[] = $criteria->compare( '==', 'text.status', 1 );
		$expr[] = $criteria->compare( '~=', 'text.content', 'ABCD:' );
		$criteria->setConditions( $criteria->combine( '&&', $expr ) );

		$textItems = $textManager->searchItems( $criteria );

		$textIds = array();
		foreach( $textItems as $item )
		{
			$textManager->deleteItem( $item->getId() );
			$textIds[] = $item->getId();
		}


		$productManager = MShop_Product_Manager_Factory::createManager( $this->_context );
		$listManager = $productManager->getSubManager( 'list' );
		$criteria = $listManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'product.list.domain', 'text' );
		$expr[] = $criteria->compare( '==', 'product.list.refid', $textIds );
		$criteria->setConditions( $criteria->combine( '&&', $expr ) );

		$listItems = $listManager->searchItems( $criteria );

		foreach( $listItems as $item ) {
			$listManager->deleteItem( $item->getId() );
		}


		foreach( $textItems as $item ) {
			$this->assertEquals( 'ABCD:', substr( $item->getContent(), 0, 5 ) );
		}

		$this->assertEquals( 6, count( $textItems ) );
		$this->assertEquals( 6, count( $listItems ) );

		if( file_exists( $filename ) !== false ) {
			throw new Exception( 'Import file was not removed' );
		}
	}
}