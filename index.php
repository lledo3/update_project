<?php


include 'excel_reader2.php';

//getting current directory
$path = './';
$dh = $path;

//Getting all the files in the current directory and putting them in a array
$files = scandir($dh);
$myFiles = array();

//Get the string position starting with "Re" and assign them to another array
	foreach($files as $sheet){
		if(strpos($sheet, "Re")===0){
			$myFiles[] = $sheet;
			}
		}
//Checks for "," or "." 
$remove_symbol = array (",",".");

//Loop through each file and grab the information needed
foreach($myFiles as $worksheet){
	
	$data = new Spreadsheet_Excel_Reader($worksheet);
	$name = $data->val(4,'C');
	$date = date("m/d/y", strtotime($data->val(6,'C')));
	$passback = str_replace($remove_symbol, "", $data->raw(11,'F'));
	
	$totalImp = $passback + str_replace($remove_symbol, "", $data->raw(12,'F')) + str_replace($remove_symbol, "", $data->raw(13,'F')) + str_replace($remove_symbol, "", $data->raw(14,'F')) + str_replace($remove_symbol, "", $data->raw(15,'F')) + str_replace($remove_symbol, "", $data->raw(16,'F'));
	
	$fillImp = $totalImp - $passback;
	
	$rev = ($data->raw(12,'J') + $data->raw(13,'J') + $data->raw(14,'J') + $data->raw(15,'J') + $data->raw(16,'J'))/2;
	$ecpm = ($rev / ($totalImp/1000));
	
	$earnings = round($rev, 4);
	$total_ecpm = round($ecpm, 4);
	
?>

<!--Output the results to the table -->
	<table border="1">
		<tbody>
			<tr>
				<td align="center"></td>
				<td align="center"><?php echo $date; ?></td>
				<td align="center"><?php echo $name; ?></td>
				<td align="center"></td>
			</tr>

			<tr>
				<td align="center"></td>
				<td align="center">Impressions</td>
				<td align="center">Earnings</td>
				<td align="center">eCPM</td>
			</tr>

			<tr>
				<td align="center">Total</td>
				<td align="center"><?php echo $totalImp; ?></td>
				<td align="center"></td>
				<td align="center"></td>
			</tr>

			<tr>
				<td align="center">Filled</td>
				<td align="center"><?php echo $fillImp; ?></td>
				<td align="center">$<?php echo $earnings; ?></td>
				<td align="center">$<?php echo $total_ecpm; ?></td>
			</tr>
		</tbody>
	</table>
	<br/><br/>

	
<?php } ?>

