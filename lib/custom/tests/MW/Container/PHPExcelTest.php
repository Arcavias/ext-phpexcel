<?php

/**
 * Test class for MW_Container_PHPExce.
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
		$container->close();
	}


	public function testNewFile()
	{
		$filename = dirname( __FILE__ ) . DIRECTORY_SEPARATOR . 'tempfile.xls';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );
		$container->close();

		$result = file_exists( $filename );
		unlink( $filename );

		$this->assertTrue( $result );
	}


	public function testAdd()
	{
		$filename = dirname( __FILE__ ) . DIRECTORY_SEPARATOR . 'tempfile.xls';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );
		$container->add( $container->create( 'test' ) );

		$result = 0;
		foreach( $container as $content ) {
			$result++;
		}

		$container->close();
		unlink( $filename );

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

		$container->close();

		$this->assertEquals( 3, $result );
	}

}
