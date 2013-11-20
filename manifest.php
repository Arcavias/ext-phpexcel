<?php

/**
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2013
 */

return array(
	'name' => 'phpexcel',
	'depends' => array(
		'arcavias-core',
	),
	'include' => array(
		'lib/custom/src',
		'lib/phpexcel',
		'controller/extjs/src',
	),
	'config' => array(
	),
	'setup' => array(
	),
	'custom' => array(
		'controller/extjs' => array(
			'controller/extjs/src',
		),
	),
);
