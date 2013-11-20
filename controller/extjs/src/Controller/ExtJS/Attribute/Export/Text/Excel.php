<?php

/**
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2011
 * @license LGPLv3, http://www.arcavias.com/en/license
 * @package Controller
 * @subpackage ExtJS
 */



/**
 * ExtJS attribute text export controller for admin interfaces.
 *
 * @package Controller
 * @subpackage ExtJS
 */
class Controller_ExtJS_Attribute_Export_Text_Excel
	extends Controller_ExtJS_Attribute_Export_Text_Default
{
	/**
	 * Creates a XLS file with all attribute texts and outputs it directly.
	 *
	 * @param stdClass $params Object containing the properties, e.g. the list of attribute IDs
	 */
	public function createHttpOutput( stdClass $params )
	{
		$this->_checkParams( $params, array( 'site', 'items' ) );
		$this->_setLocale( $params->site );

		$items = ( !is_array( $params->items ) ? array( $params->items ) : $params->items );
		$lang = ( property_exists( $params, 'lang' ) && is_array( $params->lang ) ? $params->lang : array() );

		$this->_getContext()->getLogger()->log( sprintf( 'Create export for attribute IDs: %1$s', implode( ',', $items ) ), MW_Logger_Abstract::DEBUG );


		@header('Content-Type: application/vnd.ms-excel');
		@header('Content-Disposition: attachment; filename=arcavias-attribute-texts.xls');
		@header('Cache-Control: max-age=0');

		try
		{
			$this->_exportAttributeData( $items, $lang, 'php://output' );
		}
		catch ( Exception $e )
		{
			$this->_removeTempFiles( 'php://output' );
			throw $e;
		}

		$this->_removeTempFiles( 'php://output' );
	}


	/**
	 * Creates a new job to export an excel file.
	 *
	 * @param stdClass $params Object containing the properties, e.g. the list of attribute IDs
	 */
	public function createJob( stdClass $params )
	{
		$this->_checkParams( $params, array( 'site', 'items' ) );
		$this->_setLocale( $params->site );

		$config = $this->_getContext()->getConfig();
		$dir = $config->get( 'controller/extjs/attribute/export/text/excel/exportdir', 'uploads' );

		$items = (array) $params->items;
		$lang = ( property_exists( $params, 'lang' ) ) ? (array) $params->lang : array();

		$languages = ( count( $lang ) > 0 ) ? implode( $lang, '-' ) : 'all';

		$result = (object) array(
			'site' => $params->site,
			'items' => array(
				(object) array(
					'job.label' => 'Attribute text export: '. $languages,
					'job.method' => 'Attribute_Export_Text.exportFile',
					'job.parameter' => array(
						'site' => $params->site,
						'items' => $items,
						'lang' => $params->lang,
					),
					'job.status' => 1,
				),
			),
		);

		$jobController = Controller_ExtJS_Admin_Job_Factory::createController( $this->_getContext() );
		$jobController->saveItems( $result );

		return array(
			'items' => $items,
			'success' => true,
		);
	}


	/**
	 * Create an excel file in the filesystem.
	 *
	 * @param stdClass $params Object containing the properties, e.g. the list of attribute IDs
	 */
	public function exportFile( stdClass $params )
	{
		$this->_checkParams( $params, array( 'site', 'items' ) );
		$this->_setLocale( $params->site );
		$actualLangid = $this->_getContext()->getLocale()->getLanguageId();

		$items = (array) $params->items;
		$lang = ( property_exists( $params, 'lang' ) ) ? (array) $params->lang : array();

		$config = $this->_getContext()->getConfig();
		$dir = $config->get( 'controller/extjs/attribute/export/text/excel/exportdir', 'uploads' );
		$perms = $config->get( 'controller/extjs/attribute/export/text/excel/dirperms', 0775 );

		if( is_dir( $dir ) === false && mkdir( $dir, $perms, true ) === false ) {
			throw new Controller_ExtJS_Exception( sprintf( 'Couldn\'t create directory "%1$s" with permissions "%2$o"', $dir, $perms ) );
		}

		$filename = 'attribute-text-export_' .date('Y-m-d') . '_' . md5( time() . getmypid() ) . '.xls';
		$this->_filepath = $dir . DIRECTORY_SEPARATOR . $filename;

		$this->_getContext()->getLogger()->log( sprintf( 'Create export file for attribute IDs: %1$s', implode( ',', $items ) ), MW_Logger_Abstract::DEBUG );

		try
		{
			$filename = $this->_exportAttributeData( $items, $lang, $this->_filepath );

			$this->_getContext()->getLocale()->setLanguageId( $actualLangid );
		}
		catch ( Exception $e )
		{
			$this->_removeTempFiles( $this->_filepath );
			throw $e;
		}

		return array(
			'file' => '<a href="'.$this->_filepath.'">Download</a>',
		);
	}


	/**
	 * Returns the service description of the class.
	 * It describes the class methods and its parameters including their types
	 *
	 * @return array Associative list of class/method names, their parameters and types
	 */
	public function getServiceDescription()
	{
		return array(
			'Attribute_Export_Text.createHttpOutput' => array(
				"parameters" => array(
					array( "type" => "string","name" => "site","optional" => false ),
					array( "type" => "array","name" => "items","optional" => false ),
					array( "type" => "array","name" => "lang","optional" => true ),
				),
				"returns" => "",
			),
		);
	}


	/**
	 * Inits container for storing export files.
	 *
	 * @param string $resource Path to the file
	 * @return MW_Container_Interface Container item
	 */
	protected function _initContainer( $resource )
	{
		return new MW_Container_PHPExcel( $resource, 'Excel5' );
	}
}