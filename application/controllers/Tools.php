<?php
class Tools extends CI_Controller {

     public function message($to = 'World')
     {
               echo "Hello {$to}!".PHP_EOL;
     }

     public function loop(){
     	for($i=1;$i<=100;$i++){
     		print "*";
     	}
     }

     public function test(){
     	echo base_url();
     }

}