<?php
	error_reporting(E_ALL);
?>
<html>
<head>
<title>PHP Test Page</title>
<body>
<?php
		if(isset($HTTP_RAW_POST_DATA)){
				var_dump($HTTP_RAW_POST_DATA);
		}
	echo "<br>";
	echo  $_SERVER['REQUEST_METHOD'];
	echo "<br>";
		echo  __DIR__;
	echo "<br>";
			//ini_set('max_file_uploads','10');
			
				
				
			if( isset($_REQUEST['route']) &&  $_REQUEST['route']=='fileupload'){
						var_dump($_FILES);
					
					$uploads_dir = 'uploads';
					foreach ($_FILES["userfile"]["error"] as $key => $error) {
						if ($error == UPLOAD_ERR_OK) {
							$tmp_name = $_FILES["userfile"]["tmp_name"][$key];
							// basename() may prevent filesystem traversal attacks;
							// further validation/sanitation of the filename may be appropriate
							$name = basename($_FILES["userfile"]["name"][$key]);
							move_uploaded_file($tmp_name, "$uploads_dir/$name");
						}
					}	 
				}

	
	
	var_dump($_POST);
	echo '<br>';
	var_dump($_SERVER);
				echo '<br>';
				
	echo "<br>";
	echo "hi\r\n";
	echo "<br>";

print "<img src=\"cloudoux logo.png\" width=100 height=100><h1>This Site is hosted on Hyper X 5.0</h1>";

				function isSSL() { return (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off') || $_SERVER['SERVER_PORT'] == 443; }
		//$url=$_SERVER['REQUEST_URI'];
		//$url=$_SERVER['QUERY_STRING'];
		//print_r($_SERVER);
		
		//echo $_SERVER['REQUEST_URI'];		
		$url=(isSSL()?'https://': 'http://') . $_SERVER['HTTP_HOST'].$_SERVER['REQUEST_URI'];
		$url=str_replace('?'.$_SERVER['QUERY_STRING'],'',$url);
			
				
?>
If you can see the large text and image above then php works!!!.

	<form action="<?php echo $url.'?route=fileupload'; ?>" enctype="multipart/form-data"  method="POST"  >
        <input type="hidden" name="field1" value="form_data" /><br />
        <input  multiple  name="userfile[]" type="file" /><br />
            <script type="text/javascript">
               // alert(document.forms[0]['field1'].value);
            </script>
            
        <input type="button" onclick="document.forms[0].submit();" value="submit" />
   </form>
   <hr>
	<form action="<?php echo $url.'?route=post'; ?>" enctype="application/x-www-form-urlencoded"  method="POST"  >
        <input type="hidden" name="field2" value="form_data" /><br />
        username<input  type="text"  name="username"  /><br />
            
        <input type="submit"  value="submit" />
   </form>
</body>
</html>
