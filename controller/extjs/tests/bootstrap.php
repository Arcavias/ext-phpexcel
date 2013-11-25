<?php

/**
 * @copyright Copyright (c) Metaways Infosystems GmbH, 2013
 * @version $Id$
 */


/*
 * Set error reporting to maximum
 */
error_reporting( -1 );


/*
 * Set include path for tests
 */
define('PATH_TESTS', dirname( __FILE__ ));

require_once 'TestHelper.php';
TestHelper::bootstrap();
