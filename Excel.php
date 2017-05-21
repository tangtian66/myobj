<?php 
	defined('BASEPATH') OR exit('No direct script access allowed');
	class Excel extends CI_Controller{
		
	function __construct(){
		parent::__construct();
		$this->load->helper(array('form', 'url'));
		$this->load->library('PHPExcel');
		$this->load->library('PHPExcel/IOFactory');
		//$signPackage=$this->jssdk->signPackage;
		
		//print_r($signPackage);
	}
		
	public function index(){
		
		$this->load->view("admin/excel");
		
	}	
		
	public function do_upload(){
        $config['upload_path']      = './uploads/excel/';
		$config['file_name']      = time().rand(10, 99);
        $config['allowed_types']    = 'xlsx|xls';
        $config['max_size']     = 10000;
   		$filen='';

        $this->load->library('upload', $config);

        if ( ! $this->upload->do_upload('excel'))
        {
            $error = array('error' => $this->upload->display_errors());
			print_r($error);
			exit;
            //$this->load->view('upload_form', $error);
        }
        else
        {
            $data = array('upload_data' => $this->upload->data());
			$file['file_name']=$data['upload_data']['file_name'];
			if($filen=='file'){
				echo json_encode($file);
				exit;
			}else{
				$filePath='./uploads/excel/'.$file['file_name'];		
				$data=$this->format_excel2array($filePath,$sheet=0);//./uploads/excel/'.$file['file_name']
				//print_r($data);
			}
            
        }
    }
	
	function format_excel2array($filePath='',$sheet=0){
        if(empty($filePath) or !file_exists($filePath)){die('file not exists');}
        $PHPReader = new PHPExcel_Reader_Excel2007();        //建立reader对象
        if(!$PHPReader->canRead($filePath)){
                $PHPReader = new PHPExcel_Reader_Excel5();
                if(!$PHPReader->canRead($filePath)){
                        echo 'no Excel';
                        return ;
                }
        }
        $PHPExcel = $PHPReader->load($filePath);        //建立excel对象
        $currentSheet = $PHPExcel->getSheet($sheet);        //**读取excel文件中的指定工作表*/
        $allColumn = $currentSheet->getHighestColumn();        //**取得最大的列号*/
        $allRow = $currentSheet->getHighestRow();        //**取得一共有多少行*/
        $data = array();
        for($rowIndex=1;$rowIndex<=$allRow;$rowIndex++){        //循环读取每个单元格的内容。注意行从1开始，列从A开始
                for($colIndex='A';$colIndex<=$allColumn;$colIndex++){
                        $addr = $colIndex.$rowIndex;
                        $cell = $currentSheet->getCell($addr)->getValue();
                        if($cell instanceof PHPExcel_RichText){ //富文本转换字符串
                                $cell = $cell->__toString();
                        }
                        $data[$rowIndex][$colIndex] = $cell;
                }
        }
        return $data;
	}
		
		
	}
	
	
	?>
