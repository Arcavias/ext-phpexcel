<?php

/**
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2011
 * @license LGPLv3, http://www.arcavias.com/en/license
 * @package Controller
 * @subpackage ExtJS
 */



/**
 * ExtJS catalog text export controller for admin interfaces.
 *
 * @package Controller
 * @subpackage ExtJS
 */
class Controller_ExtJS_Catalog_Export_Text_Excel
	extends Controller_ExtJS_Common_Load_Text_Abstract
	implements Controller_ExtJS_Common_Load_Text_Interface
{
	private $_sheetLine = 1;


	/**
	 * Initializes the controller.
	 *
	 * @param MShop_Context_Item_Interface $context MShop context object
	 */
	public function __construct( MShop_Context_Item_Interface $context )
	{
		parent::__construct( $context, 'Catalog_Export_Text' );
	}


	/**
	 * Creates a XLS file with all catalog texts and outputs it directly.
	 *
	 * @param stdClass $params Object containing the properties, e.g. the list of catalog node IDs
	 */
	public function createHttpOutput( stdClass $params )
	{
		$this->_checkParams( $params, array( 'site', 'items' ) );
		$this->_setLocale( $params->site );
		$actualLangid = $this->_getContext()->getLocale()->getLanguageId();

		$items = ( !is_array( $params->items ) ? array( $params->items ) : $params->items );
		$lang = ( property_exists( $params, 'lang' ) && is_array( $params->lang ) ? $params->lang : array() );

		$this->_getContext()->getLogger()->log( sprintf( 'Create export for catalog IDs: %1$s', implode( ',', $items ) ), MW_Logger_Abstract::DEBUG );


		@header('Content-Type: application/vnd.ms-excel');
		@header('Content-Disposition: attachment; filename=arcavias-catalog-texts.xls');
		@header('Cache-Control: max-age=0');


		try
		{
			$this->_getContext()->getLocale()->setLanguageId( $actualLangid );

			$this->_exportCatalogData( $items, $lang, 'php://output' );
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
	 * @param stdClass $params Object containing the properties, e.g. the list of catalog IDs
	 */
	public function createJob( stdClass $params )
	{
		$this->_checkParams( $params, array( 'site', 'items' ) );
		$this->_setLocale( $params->site );

		$config = $this->_getContext()->getConfig();
		$dir = $config->get( 'controller/extjs/catalog/export/text/excel/exportdir', 'uploads' );

		$items = (array) $params->items;
		$lang = ( property_exists( $params, 'lang' ) ) ? (array) $params->lang : array();

		$languages = ( count( $lang ) > 0 ) ? implode( $lang, '-' ) : 'all';

		$result = (object) array(
			'site' => $params->site,
			'items' => array(
				(object) array(
					'job.label' => 'Catalog text export: '. $languages,
					'job.method' => 'Catalog_Export_Text.exportFile',
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
	 * @param stdClass $params Object containing the properties, e.g. the list of catalog IDs
	 */
	public function exportFile( stdClass $params )
	{
		$this->_checkParams( $params, array( 'site', 'items' ) );
		$this->_setLocale( $params->site );
		$actualLangid = $this->_getContext()->getLocale()->getLanguageId();


		$items = (array) $params->items;
		$lang = ( property_exists( $params, 'lang' ) ) ? (array) $params->lang : array();

		$config = $this->_getContext()->getConfig();
		$dir = $config->get( 'controller/extjs/catalog/export/text/excel/exportdir', 'uploads' );
		$perms = $config->get( 'controller/extjs/catalog/export/text/excel/dirperms', 0775 );

		if( is_dir( $dir ) === false && mkdir( $dir, $perms, true ) === false ) {
			throw new Controller_ExtJS_Exception( sprintf( 'Couldn\'t create directory "%1$s" with permissions "%2$o"', $dir, $perms ) );
		}

		$filename = 'catalog-text-export_' .date('Y-m-d') . '_' . md5( time() . getmypid() ) .'.xls';

		$this->_getContext()->getLogger()->log( sprintf( 'Create export file for catalog IDs: %1$s', implode( ',', $items ) ), MW_Logger_Abstract::DEBUG );

		try
		{
			$this->_exportCatalogData( $items, $lang, $filename );

			$this->_getContext()->getLocale()->setLanguageId( $actualLangid );
		}
		catch ( Exception $e )
		{
			$this->_removeTempFiles( $tmpfolder );
			throw $e;
		}

		return array(
			'file' => '<a href="'.$filename.'">Download</a>',
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
			'Catalog_Export_Text.createHttpOutput' => array(
				"parameters" => array(
					array( "type" => "string","name" => "site","optional" => false ),
					array( "type" => "array","name" => "items","optional" => false ),
					array( "type" => "array","name" => "lang","optional" => true ),
				),
				"returns" => "",
			),
		);
	}


	protected function _exportCatalogData( array $ids, array $lang, $filename, $contentFormat = '' )
	{
		$manager = MShop_Locale_Manager_Factory::createManager( $this->_getContext() );
		$globalLanguageManager = $manager->getSubManager( 'language' );

		$search = $globalLanguageManager->createSearch();
		$search->setSortations( array( $search->sort( '+', 'locale.language.id' ) ) );

		if( !empty( $lang ) ) {
			$search->setConditions( $search->compare( '==', 'locale.language.id', $lang ) );
		}

		$containerItem = $this->_initContainer( $filename );

		$start = 0;

		do
		{
			$result = $globalLanguageManager->searchItems( $search );

			foreach ( $result as $item )
			{
				$langid = $item->getId();

				$contentItem = $containerItem->create( $langid . $contentFormat  );
				$contentItem->add( array( 'Language ID', 'Catalog label', 'Catalog ID', 'List type', 'Text type', 'Text ID', 'Text' ) );
// 				$this->_getContext()->getLocale()->setLanguageId( $langid );
				$this->_addLanguage( $langid, $ids, $contentItem );
				$containerItem->add( $contentItem );
			}

			$count = count( $result );
			$start += $count;
			$search->setSlice( $start );
		}
		while( $count == $search->getSliceSize() );

		$containerItem->close();

		return $filename;
	}


	/**
	 * Adds data for the given language.
	 *
	 * @param string $langid Language id
	 * @param array $items List of of item ids whose texts should be added
	 * @param MW_Container_Content_Interface $contentItem Content item
	 */
	protected function _addLanguage( $langid, array $ids, MW_Container_Content_Interface $contentItem )
	{
		$manager = MShop_Catalog_Manager_Factory::createManager( $this->_getContext() );

		foreach( $ids as $id )
		{
			foreach( $this->_getNodeList( $manager->getTree( $id, array('text') ) ) as $item ) {
				$this->_addItem( $langid, $item, $contentItem );
			}
		}
	}


	/**
	 * Adds all texts belonging to an product item.
	 *
	 * @param string $langid Language id
	 * @param MShop_Product_Item_Interface $item product item object
	 * @param MW_Container_Content_Interface $contentItem Content item
	 */
	protected function _addItem( $langid, MShop_Catalog_Item_Interface $item, MW_Container_Content_Interface $contentItem )
	{
		$listTypes = array();
		foreach( $item->getListItems( 'text' ) as $listItem ) {
			$listTypes[ $listItem->getRefId() ] = $listItem->getType();
		}



		foreach( $this->_getTextTypes( 'catalog' ) as $textTypeItem )
		{
			$textItems = $item->getRefItems( 'text', $textTypeItem->getCode() );

			if( !empty( $textItems ) )
			{
				foreach( $textItems as $textItem )
				{
					$listType = ( isset( $listTypes[ $textItem->getId() ] ) ? $listTypes[ $textItem->getId() ] : '' );

					$items = array( $langid, $item->getLabel(), $item->getId(), $listType, $textTypeItem->getCode(), '', '' );

					// use language of the text item because it may be null
					if( ( $textItem->getLanguageId() == $langid || is_null( $textItem->getLanguageId() ) )
						&& $textItem->getTypeId() == $textTypeItem->getId() )
					{
						$items[0] = $textItem->getLanguageId();
						$items[5] = $textItem->getId();
						$items[6] = $textItem->getContent();
					}
				}
			}
			else
			{
				$items = array( $langid, $item->getLabel(), $item->getId(), 'default', $textTypeItem->getCode(), '', '' );
			}

			$contentItem->add( $items );
		}
	}


	/**
	 * Inits container for storing export files.
	 *
	 * @param string $resource Path or resource
	 * @return MW_Container_Interface Container item
	 */
	protected function _initContainer( $resource )
	{
		return new MW_Container_PHPExcel( $resource, 'Excel5' );
	}


	/**
	 * Get all child nodes.
	 *
	 * @param MShop_Catalog_Item_Interface $node
	 * @return array $nodes List of nodes
	 */
	protected function _getNodeList( MShop_Catalog_Item_Interface $node )
	{
		$nodes = array( $node );

		foreach( $node->getChildren() as $child ) {
			$nodes = array_merge( $nodes, $this->_getNodeList( $child ) );
		}

		return $nodes;
	}
}
