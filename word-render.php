<?php
/**
 * @author: msztorc
 * @date:   2016-06-20 12:27:10
 * @last modified by:   msztorc
 * @last modified time: 2016-07-18 12:31:45
 */

$cache_dir = '.cache';

$shortopts  = "";
$shortopts .= "v";  	// verbose mode
$shortopts .= "h";  	// help
$shortopts .= "i:"; 	// input docx file
$shortopts .= "o:"; 	// output file (pdf or docx)
$shortopts .= "r:"; 	// replace patterns and values


$longopts  = array(
    "verbose",  		// verbose mode
    "help",     		// help
    "input:",   		// input docx file
    "output:",  		// output file (pdf or docx)
    "replace:", 		// replace patterns and values

);
$options = getopt($shortopts, $longopts);


if (isset($options['h']) || isset($options['help']) || (!isset($options['i']) && !isset($options['input'])) || (!isset($options['o']) && !isset($options['output'])))
{

    echo "\n\033[32mWordRender\033[0m v1.0\n\n";
    echo "Usage: php word-render.php [options]\n";

    echo "Options:\n";
    echo "-v, --verbose \t\tVerbose mode\n";
    echo "-h, --help \t\tHelp\n";
    echo "-i, --input \t\tInput file [required]\n";
    echo "-o, --output \t\tOutput file (docx, pdf) [required]\n";   
    echo "-r, --replace \t\tPlaceholders with values (placeholder1=>value|placeholder2=>value...) or json file\n";  
    die();
}

$verbose = (isset($options['v']) || isset($options['verbose'])) ? true : false;


function rmdir_recursive($dir) { 
    $files = array_diff(scandir($dir), array('.','..')); 
    foreach ($files as $file) { 
      (is_dir($dir . DIRECTORY_SEPARATOR . $file)) ? rmdir_recursive($dir . DIRECTORY_SEPARATOR . $file) : unlink($dir . DIRECTORY_SEPARATOR . $file); 
    } 
    return rmdir($dir); 
}


$file_input = (isset($options['input'])) ? $options['input'] : '';
if ($file_input == '' && isset($options['i'])) $file_input = $options['i'];

$file_output = (isset($options['output'])) ? $options['output'] : '';
if ($file_output == '' && isset($options['o'])) $file_output = $options['o'];    

if (!file_exists($file_input)) die("Input file doesn't exists! ($file_input)\n");

$replace = (isset($options['replace'])) ? $options['replace'] : '';
if ($replace == '' && isset($options['r'])) $replace = $options['r'];

//replace
if ($replace != '') {

    if ($verbose) echo "\nExtracting input file...";

	$zip = new ZipArchive();
	$handle = $zip->open($file_input);

	if ($handle === true) {
		$temp_dir = $cache_dir . DIRECTORY_SEPARATOR . substr(md5(microtime()), -6);

		$zip->extractTo($temp_dir);
		$zip->close();


	} else die("Can't read input file\n");

    if ($verbose) echo "[ok]\n";

	$replacements = (file_exists($replace)) ? json_decode(file_get_contents($replace), true) : explode('|', $replace);

	if (count($replacements) > 0) {

		$template = file_get_contents($temp_dir . DIRECTORY_SEPARATOR . 'word' . DIRECTORY_SEPARATOR . 'document.xml');

		$patterns = [];
		$values = [];

        if ($verbose) echo "\nApplying replacements...\n";

		foreach($replacements as $replacement => $val) {
			
            list($placeholder, $value) = (file_exists($replace)) ? [$replacement, $val] : explode('=>', $val);

            if ($verbose) echo $placeholder ." => " . $value . "\n";

            $patterns[] = '/'. $placeholder .'/';
            //$patterns[] = '/{(.+?)'. $placeholder .'(.+?)}/';
			//$patterns[] = '/{[^_]*'. $placeholder .'[^_}]*}/';
			$values[] = $value;
		}

        $document = preg_replace($patterns, $values, $template);
		//$document = preg_replace('/{[^_]*(_\w+)+[^_}]*}/', '', preg_replace($patterns, $values, $template));

		file_put_contents($temp_dir . DIRECTORY_SEPARATOR . 'word' . DIRECTORY_SEPARATOR . 'document.xml', $document);

        try {
            if ($verbose) echo "\nSaving document...";

            $zip = new ZipArchive();

            $file_input = pathinfo($file_output, PATHINFO_DIRNAME) . DIRECTORY_SEPARATOR . pathinfo($file_output, PATHINFO_FILENAME) . '.docx';

            if ($zip->open($file_input, ZIPARCHIVE::CREATE ) !== true ){
                die('Error creating output filename [zip]');
            }

            $iterator = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($temp_dir . DIRECTORY_SEPARATOR, FilesystemIterator::SKIP_DOTS));

            foreach ($iterator as $key=>$value) {
                $localpath=substr($key, strlen($temp_dir.DIRECTORY_SEPARATOR));
                $zip->addFile(realpath($key), $localpath) or die ("Could not add file [zip]: $key");
            }
            $zip->close();  
            if ($verbose) echo "[ok]\n";  

        } catch (Exception $e) {
            if ($verbose) echo "[fail]\n"; 
            print_r($e);                
        }

        rmdir_recursive($temp_dir);        

	}
}

// only for Windows with MS Word installed
if (strtolower(pathinfo($file_output, PATHINFO_EXTENSION)) == 'pdf') {

    if (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN') {

        try {
            if ($verbose) echo "\nExporting to PDF ($file_output)...";

            $word = new COM('Word.Application') or die('Could not open MS Word Application');

            $word->Documents->Open(realpath($file_input));

            $word->ActiveDocument->ExportAsFixedFormat(realpath(dirname($file_output)) . DIRECTORY_SEPARATOR . basename($file_output), 17, false, 0, 0, 0, 0, 7, true, true, 2, true, true, false);
            $word->ActiveDocument->Close(false);
            $word->Quit(false);   
            if ($verbose) echo "[ok]\n"; 

        } catch (Exception $e) {
            if ($verbose) echo "[fail]\n"; 

            print_r($e);
            $word->Quit(false);
        } 

    } else die("\nRun this script under Windows with MS Word installed.\n");
}