<?php defined('BASEPATH') OR exit('No direct script access allowed');

class Export extends CI_Controller{

	var $file  =  'template.xlsx';

	public function __Construct(){
		parent::__Construct();
		$this->load->database();
		require_once("vendor/phpoffice/phpexcel/Classes/PHPExcel.php");
	}

	public function get($table,$param,$value){
		$this->db->where('param',$value);
		$this->db->limit(1);
		$result = $this->db->get($table)->row();
		return $result->$param;
	}

	public function test(){
		$file = "files/".$this->file;
		$excel = PHPExcel_IOFactory::load($file);
		$cell_collection = $excel->getActiveSheet()->getCellCollection();
		$data = array();
		$start = 5;
		$max = $excel->getActiveSheet()->getHighestRow();
		$time_start = microtime(true); 

		$date_planned = $excel->getSheet(0)->getCell('K'.$start)->getValue();

		$this->db->where('assignment_id',10017);
		$this->db->join('sow_detail','sow_detail.sow_detail_id = assignment.sow_detail_id');
		$this->db->join('sow','sow.sow_id = sow_detail.sow_id');
		$asg = $this->db->get('assignment')->row();


		
		   
	}

	public function run(){

		echo "Please Wait...."
		$file = "files/".$this->file;
		$excel = PHPExcel_IOFactory::load($file);
		$cell_collection = $excel->getActiveSheet()->getCellCollection();
		$data = array();
		$start = 2;
		$max = $excel->getActiveSheet()->getHighestRow();
		$time_start = microtime(true); 
		$type = null;
		
		$number = 0;
		for($i=$start;$i<=$max;$i++){

			$assignment_id = $excel->getSheet(0)->getCell('A'.$i)->getValue();
			$activity_id = $excel->getSheet(0)->getCell('B'.$i)->getValue();
			$project_name = $excel->getSheet(0)->getCell('C'.$i)->getValue();
			$report_type = $excel->getSheet(0)->getCell('D'.$i)->getValue();
			$region_name = $excel->getSheet(0)->getCell('E'.$i)->getValue();
			$site_id = $excel->getSheet(0)->getCell('F'.$i)->getValue();
			$site_name = $excel->getSheet(0)->getCell('G'.$i)->getValue();
			$latitude = $excel->getSheet(0)->getCell('H'.$i)->getValue();
			$longitude = $excel->getSheet(0)->getCell('I'.$i)->getValue();
			$vendor = $excel->getSheet(0)->getCell('J'.$i)->getValue();
			$date_planned = date($format = "Y-m-d H:i:s", PHPExcel_Shared_Date::ExcelToPHP($excel->getSheet(0)
				->getCell('K'.$i)->getValue()));
			$date_completion = $excel->getSheet(0)->getCell('L'.$i)->getValue();
			$batch = $excel->getSheet(0)->getCell('M'.$i)->getValue();
			$status = $excel->getSheet(0)->getCell('N'.$i)->getValue();


			// REPORT TYPE / TECHNOLOGY PROCESS
			$technology_process_id = null;
			$this->db->where('technology_process',$report_type);
			$this->db->limit(1);
			$technology_process = $this->db->get('technology_process')->row();
			if($technology_process){
				$technology_process_id = $technology_process->id;
			}else{
				$this->db->insert('technology_process',array('technology_process'=>$report_type));
				$technology_process_id = $this->db->insert_id();
			}


			// REGION
			$region_id = null;
			$this->db->where('region_name',$region_name);
			$this->db->limit(1);
			$region = $this->db->get('region')->row();
			if($region){
				$region_id = $region->region_id;
			}else{
				$this->db->insert('region',array('region_name'=>$region_name));
				$region_id = $this->db->insert_id();
			}

			// SITE
			$sites_id = null;
			$this->db->where('site_id',$site_id);
			$site = $this->db->get('site')->row();
			if($site){
				$sites_id = $site->site_id;
				$this->db->where('site_id',$site_id);
				$this->db->update('site',array(
					'site_id'=>$site_id,
					'site_name'=>$site_name,
					'lat'=>$latitude,
					'lon'=>$longitude
				));
			}else{
				$this->db->insert('site',array(
					'site_id'=>$site_id,
					'site_name'=>$site_name,
					'lat'=>$latitude,
					'lon'=>$longitude
				));
				$sites_id = $this->db->insert_id();
			}

			
			// VENDOR
			$asp_id = null;
			$this->db->where('asp_name',$vendor);
			$this->db->limit(1);
			$asp = $this->db->get('asp')->row();
			if($asp){
				$asp_id = $asp->asp_id;
			}else{
				$this->db->insert('asp',array('asp_name'=>$vendor));
				$asp_id = $this->db->insert_id();
			}

			// STATUS
			$assignment_status_id = null;
			$this->db->where('assignment_status_desc',strtoupper($status));
			$this->db->limit(1);
			$assignment_status = $this->db->get('assignment_status')->row();
			if($assignment_status){
				$assignment_status_id = $assignment_status->assignment_status_id;
			}else{
				$this->db->insert('assignment_status',array('assignment_status_desc'=>strtoupper($status)));
				$assignment_status_id = $this->db->insert_id();
			}

			
			$this->db->where('assignment_id',$assignment_id);
			$this->db->join('sow_detail','sow_detail.sow_detail_id = assignment.sow_detail_id');
			$this->db->join('sow','sow.sow_id = sow_detail.sow_id');
			$asg = $this->db->get('assignment')->row();


			if($asg){ // KONDISI ADA

				$type = 'Updated';

				// Update SOW
				$this->db->where('sow_id',$asg->sow_id);
				$this->db->update('sow',array(
					'technology_subprocess_id'=>1,
					'region_id'=>$region_id,
					'site_id'=>$site_id,
					'technology_process_id'=>$technology_process_id,
					'status'=>0
				));
				


				// Update SOW_DETAIL
				$this->db->where('sow_detail_id',$asg->sow_detail_id);
				$this->db->update('sow_detail',array(
					'sow_id'=>$asg->sow_id,
					'sow_detail_desc'=>$project_name,
					'form_type'=>0
				));
			

				$this->db->where('asp_team_id',$asg->asp_team_id);
				$this->db->update('asp_team',array(
					'asp_team_desc'=>$vendor,
					'asp_id'=>$asg->asp_id
				));
				



				// Update ASSIGNMENT
				$this->db->where('assignment_id',$assignment_id);
				$update = $this->db->update('assignment',array(
					'assignment_id'=>$assignment_id,
					'sow_detail_id'=>$asg->sow_detail_id,
					'asp_team_id'=>$asg->asp_team_id,
					'creator_id'=>'admin',
					'asp_id'=>$asp_id,
					'itc_type'=>0,
					'created_by'=>'admin',
					'current_status'=>$assignment_status_id,
					'date_planned'=>$date_planned,
					'batch'=>$batch
				));

				if($update){
					$number++;
				}

				
				
			}else{ // KONDISI KOSONG

				$type = 'Created';

				// INSERT SOW
				$this->db->insert('sow',array(
					'technology_subprocess_id'=>1,
					'region_id'=>$region_id,
					'site_id'=>$site_id,
					'technology_process_id'=>$technology_process_id,
					'status'=>0
				));
				$sow_id = $this->db->insert_id();


				// INSERT SOW_DETAIL
				$this->db->insert('sow_detail',array(
					'sow_id'=>$sow_id,
					'sow_detail_desc'=>$project_name,
					'form_type'=>0
				));
				$sow_detail_id = $this->db->insert_id();


				$this->db->insert('asp_team',array(
					'asp_team_desc'=>$vendor,
					'asp_id'=>$asp_id
				));
				$asp_team_id = $this->db->insert_id();


				// INSERT ASSIGNMENT
				$insert = $this->db->insert('assignment',array(
					'assignment_id'=>$assignment_id,
					'sow_detail_id'=>$sow_detail_id,
					'asp_team_id'=>$asp_team_id,
					'creator_id'=>'admin',
					'asp_id'=>$asp_id,
					'itc_type'=>0,
					'created_by'=>'admin',
					'current_status'=>$assignment_status_id,
					'date_planned'=>$date_planned,
					'batch'=>$batch
				));

				if($insert){
					$number++;
				}

			}
			
			
		}

		$time_end = microtime(true);
		$execution_time = ($time_end - $time_start)/60;
		echo 'Total Execution Time: '.$execution_time.' Minuite. On '.$number.' Data was '.$type;
	}

}