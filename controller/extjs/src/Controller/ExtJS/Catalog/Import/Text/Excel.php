<?php

/**
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2011
 * @license LGPLv3, http://www.arcavias.com/en/license
 * @package Controller
 * @subpackage ExtJS
 */


/**
 * ExtJS catalog text import controller for admin interfaces.
 *
 * @package Controller
 * @subpackage ExtJS
 */
class Controller_ExtJS_Catalog_Import_Text_Excel
	extends Controller_ExtJS_Catalog_Import_Text_Default
{
	/**
	 * Uploads a XLS file with all catalog texts.
	 *
	 * @param stdClass $params Object containing the properties
	 */
	public function uploadFile( stdClass $params )
	{
		$this->_checkParams( $params, array( 'site' ) );
		$this->_setLocale( $params->site );

		if( ( $fileinfo = reset( $_FILES ) ) === false ) {
			throw new Controller_ExtJS_Exception( 'No file was uploaded' );
		}

		$config = $this->_getContext()->getConfig();
		$dir = $config->get( 'controller/extjs/catalog/import/text/excel/uploaddir', 'uploads' );

		if( $config->get( 'controller/extjs/catalog/import/text/excel/enablecheck', true ) ) {
			$this->_checkFileUpload( $fileinfo['tmp_name'], $fileinfo['error'] );
		}

		$fileext = pathinfo( $fileinfo['name'], PATHINFO_EXTENSION );
		$dest = $dir . DIRECTORY_SEPARATOR . md5( $fileinfo['name'] . time() . getmypid() ) . '.' . $fileext;

		if( rename( $fileinfo['tmp_name'], $dest ) !== true )
		{
			$msg = sprintf( 'Uploaded file could not be moved to upload directory "%1$s"', $dir );
			throw new Controller_ExtJS_Exception( $msg );
		}

		$perms = $config->get( 'controller/extjs/catalog/import/text/excel/fileperms', 0660 );
		if( chmod( $dest, $perms ) !== true )
		{
			$msg = sprintf( 'Could not set permissions "%1$s" for file "%2$s"', $perms, $dest );
			throw new Controller_ExtJS_Exception( $msg );
		}

		$result = (object) array(
			'site' => $params->site,
			'items' => array(
				(object) array(
					'job.label' => 'Catalog text import: ' . $fileinfo['name'],
					'job.method' => 'Catalog_Import_Text.importFile',
					'job.parameter' => array(
						'site' => $params->site,
						'items' => $dest,
					),
					'job.status' => 1,
				),
			),
		);

		$jobController = Controller_ExtJS_Admin_Job_Factory::createController( $this->_getContext() );
		$jobController->saveItems( $result );

		return array(
			'items' => $dest,
			'success' => true,
		);
	}


	/**
	 * Imports a file that can be understood by PHPExcel.
	 *
	 * @param string $path Path to file for importing
	 */
	protected function _importFile( $path )
	{
		$container = new MW_Container_PHPExcel( $path, 'Excel5' );

		$textTypeMap = array();
		foreach( $this->_getTextTypes( 'catalog' ) as $item ) {
			$textTypeMap[ $item->getCode() ] = $item->getId();
		}

		foreach( $container as $langContent )
		{
			$catalogTextMap = $this->_importTextsFromContent( $langContent, $textTypeMap, 'catalog' );
			$this->_importCatalogReferences( $catalogTextMap );
		}
	}
}
