PHP-DotNet-Bridge
=================

A PHP <-> .NET bridge via VB.net Reflection.  Similar to PHP's DOTNET class (http://php.net/manual/en/class.dotnet.php) but more awesome.  Based on code that used to be on sourceforge (see https://www.codeproject.com/Articles/113720/Universal-COM-Callable-Wrapper and https://sourceforge.net/projects/universalccw/ )

Differences from DOTNET
-----------------------
  * Can load .net libraries that aren't in the Global Assembly Cache (and thus, don't need to be strongly-named)
  * Can instantiate objects with parameters in their constructors
  * Can modify fields in Struct/Structures (unboxing)
  * Can instantiate Struct/Structures
  * Works with .net 4
  * 

Installation
------------
Use regasm to install the dll, eg `C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe Universal_CCW.dll`

Example
-------
<code>
CCWObjectWrapper::Load_DOTNET_Dll("C:\Local\3rdparty.dll");

$thirdPartyStruct = new CCWObjectWrapper('theAssembly', 'structName');
$thirdPartyStruct->seqNo = 123;

$thirdPartyObj = new CCWObjectWrapper('theAssembly', 'theAssembly.theClass', array($param1, $param2));
echo $thirdPartyObj->doSomething($thirdPartyStruct);
$thirdPartyObj->something = 'another thing';

$thirdPartyObj = null;
</code>

Compiling
---------
To compile use `C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe /vbruntime+ /warnaserror+ /target:library Universal_CCW_source.vb /out:Universal_CCW.dll`
