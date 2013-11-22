<?php

/**
 * Test class for MW_Container_PHPExcel.
 *
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2013
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 */


class MW_Container_PHPExcelTest extends MW_Unittest_Testcase
{

	public function testExistingFile()
	{
		$filename = dirname( __FILE__ ) . DIRECTORY_SEPARATOR . 'excel5.xls';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );
	}


	public function testNewFile()
	{
		$filename = dirname( __FILE__ ) . DIRECTORY_SEPARATOR . 'tempfile';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );
		$container->close();

		$result = file_exists( $container->getName() );
		unlink( $container->getName() );

		$this->assertTrue( $result );
		$this->assertEquals( '.xls', substr( $container->getName(), -4 ) );
		$this->assertFalse( file_exists( $container->getName() ) );
	}


	public function testFormat()
	{
		$container = new MW_Container_PHPExcel( 'tempfile', 'Excel2007', array() );
		$this->assertEquals( '.xlsx', substr( $container->getName(), -5 ) );

		$container = new MW_Container_PHPExcel( 'tempfile', 'OOCalc', array() );
		$this->assertEquals( '.ods', substr( $container->getName(), -4 ) );

		$container = new MW_Container_PHPExcel( 'tempfile', 'CSV', array() );
		$this->assertEquals( '.csv', substr( $container->getName(), -4 ) );
	}


	public function testAdd()
	{
		$filename = dirname( __FILE__ ) . DIRECTORY_SEPARATOR . 'tempfile';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );
		$container->add( $container->create( 'test' ) );

		$result = 0;
		foreach( $container as $content ) {
			$result++;
		}

		$container->close();
		unlink( $container->getName() );

		$this->assertEquals( 1, $result );
	}


	public function testIterator()
	{
		$filename = dirname( __FILE__ ) . DIRECTORY_SEPARATOR . 'excel5.xls';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );

		$result = 0;
		foreach( $container as $content ) {
			$result++;
		}

		$this->assertEquals( 3, $result );
	}

}
