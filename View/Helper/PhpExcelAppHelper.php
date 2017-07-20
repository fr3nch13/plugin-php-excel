<?php

App::uses('Helper', 'View');
App::uses('AppHelper', 'View/Helper');

class PhpExcelAppHelper extends AppHelper 
{
	public $helpers = array(
		'Time',
		'Number',
	);
}
