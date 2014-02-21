<?php

class CCWObjectWrapper {
	public $COM;
	private $eventName;
	
	function __construct($comObjectOrAssembly, $className = null, $constructorArgs = null) {
		static $eventNumCount = 1;
		
		if (is_string($comObjectOrAssembly)) {
			$this->eventName = 'obj' . $eventNumCount++;
			if (empty($constructorArgs))
				$comObjectOrAssembly = CCWObjectWrapper::Universal_CCW_Factory()->New_Object($this->eventName, $comObjectOrAssembly, $className);
			else
				$comObjectOrAssembly = CCWObjectWrapper::Universal_CCW_Factory()->New_Object($this->eventName, $comObjectOrAssembly, $className, $constructorArgs);
		}
		
		$this->COM = $comObjectOrAssembly;
	}
	
	function __get($property) {
		$value = $this->COM->Get_Property_Value($property);
		if (gettype($value) == 'object')
			$value = new CCWObjectWrapper($value);
		
		return $value;
	}
	
	function __set($property, $value) {
		if (is_a($value, 'CCWObjectWrapper'))
			$value = $value->COM;
		
		$this->COM->Set_Property_Value($property, $value);
	}
	
	function __call($function, $arguments) {
		foreach($arguments as $argkey => $argvalue)
		{
			if (is_a($argvalue, 'CCWObjectWrapper'))
				$arguments[$argkey] = $argvalue->COM;
		}
		
		$value = $this->COM->Call_Method($function, $arguments);
		
		if (gettype($value) == 'object')
			$value = new CCWObjectWrapper($value);
		
		return $value;
	}
	
	static function Universal_CCW_Factory() {
		static $factory = null;
		if ($factory !== null)
			return $factory;
		$factory = new COM("Universal_CCW.Universal_CCW_Factory");
		return $factory;
	}
	
	static function Load_DOTNET_Dll($dllPath) {
		$reflection = $COM->New_Static('mscorlib', 'System.Reflection.Assembly');
		$reflection->Call_Static_Method('LoadFrom', array(realpath($dllPath)));
	}
}
