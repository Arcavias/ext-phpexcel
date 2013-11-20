<?php

/**
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2013
 * @license LGPLv3, http://www.arcavias.com/en/license
 * @package MW
 * @subpackage Container
 */


/**
 * Implementation of PHPExcel containers.
 *
 * @package MW
 * @subpackage Container
 */
class MW_Container_PHPExcel implements MW_Container_Interface
{
	private $_container;
	private $_options;
	private $_format;


	/**
	 * Opens an existing container or creates a new one.
	 *
	 * @param string $resourcepath Path to the resource like a file
	 * @param string $format Format of the content objects inside the container
	 * @param array $options Associative list of key/value pairs for configuration
	 */
	public function __construct( $resourcepath, $format, array $options = array() )
	{
		if( file_exists( $resourcepath ) )
		{
			$type = PHPExcel_IOFactory::identify( $resourcepath );
			$reader = PHPExcel_IOFactory::createReader( $type );
			$this->_container = $reader->load( $resourcepath );
		}
		else
		{
			$this->_container = new PHPExcel();
			$this->_container->removeSheetByIndex( 0 );
		}

		$this->_iterator = $this->_container->getWorksheetIterator();

		$this->_resourcepath = $resourcepath;
		$this->_options = $options;
		$this->_format = $format;
	}


	/**
	 * Creates a new content object.
	 *
	 * @param string $name Name of the content
	 * @return MW_Container_Content_Interface New content object
	 */
	public function create( $name )
	{
		$sheet = $this->_container->createSheet();
		$sheet->setTitle( $name );

		return new MW_Container_Content_PHPExcel( $sheet, $name, $this->_options );
	}


	/**
	 * Adds content data to the container.
	 *
	 * @param MW_Container_Content_Interface $content Content object
	 */
	public function add( MW_Container_Content_Interface $content )
	{
		// was already added to the PHPExcel object by createSheet()
	}


	/**
	 * Cleans up and saves the container.
	 */
	public function close()
	{
		$writer = PHPExcel_IOFactory::createWriter( $this->_container, $this->_format );
		$writer->save( $this->_resourcepath );
	}


	/**
	 * Return the current element.
	 *
	 * @return MW_Container_Content_Interface Content object with PHPExcel sheet
	 */
	function current()
	{
		$sheet = $this->_iterator->current();

		return new MW_Container_Content_PHPExcel( $sheet, $sheet->getTitle(), $this->_options );
	}


	/**
	 * Returns the key of the current element.
	 *
	 * @return integer Index of the PHPExcel sheet
	 */
	function key()
	{
		return $this->_iterator->key();
	}


	/**
	 * Moves forward to next PHPExcel sheet.
	 */
	function next()
	{
		return $this->_iterator->next();
	}


	/**
	 * Rewinds to the first PHPExcel sheet.
	 */
	function rewind()
	{
		return $this->_iterator->rewind();
	}


	/**
	 * Checks if the current position is valid.
	 *
	 * @return boolean True on success or false on failure
	 */
	function valid()
	{
		return $this->_iterator->valid();
	}
}