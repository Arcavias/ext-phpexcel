<?php

/**
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2011
 * @license LGPLv3, http://www.arcavias.com/en/license
 */


class Controller_ExtJS_Catalog_Import_Text_ExcelTest extends MW_Unittest_Testcase
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

		$suite  = new PHPUnit_Framework_TestSuite( 'Controller_ExtJS_Catalog_Import_Text_ExcelTest' );
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
		$this->_context->getConfig()->set( 'controller/extjs/catalog/export/text/default/container/type', 'PHPExcel' );
		$this->_context->getConfig()->set( 'controller/extjs/catalog/export/text/default/container/format', 'Excel5' );
		$this->_context->getConfig()->set( 'controller/extjs/catalog/import/text/default/container/type', 'PHPExcel' );
		$this->_context->getConfig()->set( 'controller/extjs/catalog/import/text/default/container/format', 'Excel5' );

		$this->_object = new Controller_ExtJS_Catalog_Import_Text_Default( $this->_context );
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
		$this->_object = new Controller_ExtJS_Catalog_Import_Text_Default( $this->_context );

		$catalogManager = MShop_Catalog_Manager_Factory::createManager( $this->_context );

		$node = $catalogManager->getTree( null, array(), MW_Tree_Manager_Abstract::LEVEL_ONE );

		$params = new stdClass();
		$params->lang = array( 'en' );
		$params->items = $node->getId();
		$params->site = $this->_context->getLocale()->getSite()->getCode();


		$exporter = new Controller_ExtJS_Catalog_Export_Text_Default( $this->_context );
		$result = $exporter->exportFile( $params );

		$this->assertTrue( array_key_exists('file', $result) );

		$filename = substr($result['file'], 9, -14);
		$this->assertTrue( file_exists( $filename ) );

		$filename2 = 'catalog-import.xls';

		$phpExcel = PHPExcel_IOFactory::load($filename);

		if( unlink( $filename ) !== true ) {
			throw new Exception( sprintf( 'Deleting file "%1$s" failed', $filename ) );
		}

		$sheet = $phpExcel->getSheet( 0 );

		$sheet->setCellValueByColumnAndRow( 6, 2, 'Root: delivery info' );
		$sheet->setCellValueByColumnAndRow( 6, 3, 'Root: long' );
		$sheet->setCellValueByColumnAndRow( 6, 4, 'Root: name' );
		$sheet->setCellValueByColumnAndRow( 6, 5, 'Root: payment info' );
		$sheet->setCellValueByColumnAndRow( 6, 6, 'Root: short' );

		$objWriter = PHPExcel_IOFactory::createWriter( $phpExcel, 'Excel5' );
		$objWriter->save( $filename2 );


		$params = new stdClass();
		$params->site = $this->_context->getLocale()->getSite()->getCode();
		$params->items = $filename2;

		$this->_object->importFile( $params );

		if( file_exists( $filename2 ) !== false ) {
			throw new Exception( 'Import file was not removed' );
		}

		$textManager = MShop_Text_Manager_Factory::createManager( $this->_context );
		$criteria = $textManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'text.domain', 'catalog' );
		$expr[] = $criteria->compare( '==', 'text.languageid', 'en' );
		$expr[] = $criteria->compare( '==', 'text.status', 1 );
		$expr[] = $criteria->compare( '~=', 'text.content', 'Root:' );
		$criteria->setConditions( $criteria->combine( '&&', $expr ) );

		$textItems = $textManager->searchItems( $criteria );

		$textIds = array();
		foreach( $textItems as $item )
		{
			$textManager->deleteItem( $item->getId() );
			$textIds[] = $item->getId();
		}


		$listManager = $catalogManager->getSubManager( 'list' );
		$criteria = $listManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'catalog.list.domain', 'text' );
		$expr[] = $criteria->compare( '==', 'catalog.list.refid', $textIds );
		$criteria->setConditions( $criteria->combine( '&&', $expr ) );

		$listItems = $listManager->searchItems( $criteria );

		foreach( $listItems as $item ) {
			$listManager->deleteItem( $item->getId() );
		}


		$this->assertEquals( 5, count( $textItems ) );
		$this->assertEquals( 5, count( $listItems ) );

		foreach( $textItems as $item ) {
			$this->assertEquals( 'Root:', substr( $item->getContent(), 0, 5 ) );
		}
	}
}