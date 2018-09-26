<?php namespace crocodicstudio\crudbooster\controllers;

error_reporting(E_ALL ^ E_NOTICE);

use crocodicstudio\crudbooster\controllers\CBController;
use Illuminate\Support\Facades\Session;
use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Facades\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\App;
use Illuminate\Support\Facades\Mail;
use Illuminate\Support\Facades\Route;
use Illuminate\Support\Facades\Hash;
use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Validator;
use Illuminate\Support\Facades\PDF;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\Log;
use CRUDBooster;
use CB;
use Schema;
use JsValidator;
use GSSDK;
use GSRequest;

class GigyaCBController extends CBController {
	
	private $gigya_api_key;
	private $gigya_secret_key;
	private $gigya_user_key;

	public function __construct()
	{
		$this->gigya_api_key  = config('crudbooster.GIGYAAPIKEY');
		$this->gigya_secret_key = config('crudbooster.GIGYASECRETKEY');
		$this->gigya_user_key = config('crudbooster.GIGYAUSERKEY');
	}

	private function getCustomer($offset=0,$limit=300000)
    {

    	$method = "accounts.search";

    	// $request = new GSRequest($apiKey,$secretKey,$method);
    	$request = new GSRequest($this->gigya_api_key,$this->gigya_secret_key,$method,null,true,$this->gigya_user_key);

    	$request->setParam("query","select * from accounts order by registered DESC START ".$offset." LIMIT ".$limit."");
    	// $request->setParam("openCursor",true);

    	$response = $request->send();

    	if($response->getErrorCode()==0)
    	    {
    	        // echo "Success";
                $response = $response->getResponseText();
                $response = json_decode($response, true);
                // return view('test', compact('response'));

    	        return $response;

    	    }
    	else
    	    {
    	        echo ("Uh-oh, we got the following error: " . $response->getErrorMessage());
    	        error_log($response->getLog());
    	    }


    }

	public function getIndex() {
		$this->cbLoader();

		$module = CRUDBooster::getCurrentModule();

		if(!CRUDBooster::isView() && $this->global_privilege==FALSE) {
			CRUDBooster::insertLog(trans('crudbooster.log_try_view',['module'=>$module->name]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans('crudbooster.denied_access'));
		}
		
		$data['table'] 	  = $this->table;
		$data['table_pk'] = CB::pk($this->table);
		$data['page_title']       = $module->name;
		$data['page_description'] = trans('crudbooster.default_module_description');
		$data['date_candidate']   = $this->date_candidate;
		$data['limit'] = $limit   = (Request::get('limit'))?Request::get('limit'):$this->limit;

		$tablePK = $data['table_pk'];
		$table_columns = CB::getTableColumns($this->table);
		$result = DB::table($this->table)->select(DB::raw($this->primary_key));

		$tableName = 'customergigya';


		if (!Request::get('page')) {
			//insert table
			//$tempTable = $this->createTempTable();
			$gigyaData = $this->getCustomer();

			// dd($gigyaData);

			$gigyaResults = $gigyaData['results'];
			$totalCount = $gigyaData['totalCount'];
			$fdataArray = range(0,$totalCount);
			print_r($fdataArray);
			//echo $totalCount."<br>";
			//echo "<pre>".print_r($gigyaResults,TRUE)."</pre><br>";
			
			DB::table($tableName)->truncate();

			foreach ($gigyaResults as $gigyaResult) {
				
				// $profile[] = $gigyaResult['profile'];
				$i = 0;
				$col = array_keys($gigyaResult['profile']);
				$colLength = sizeof($col);
				// $childData[] = $gigyaResult['data']['child'];
				// dump($profile);

				$listOfColumn = DB::select(DB::raw("SHOW COLUMNS in $tableName"));

				for ($a=0; $a < $colLength; $a++) { 

				    $colName = $col[$a];
				    // dd($colName);
				    $profile[$i][$colName] = $gigyaResult['profile'][$colName];
				    // dump($profile[$i][$colName]);
				    foreach ($listOfColumn as $listOfCol) {
				        $listCol[] = $listOfCol->Field;
				        
				    }
				    // dump($colName,$listCol);
				    if(!in_array($colName, $listCol)){
				        // dump($colName,'not exist');/
				        DB::insert(DB::raw("ALTER TABLE $tableName ADD COLUMN $colName varchar(255) NOT NULL"));
				    }
				}
				    

			    DB::table($tableName)->insert([
			        $profile[0]
			    ]);
				// dump($profile);
			    $i++;
			}
		}
		$customerData = DB::table($tableName)->get();

		//end insert table

		// $this->hook_query_index($result);
			# code...

		$alias            = array();
		$table            = $this->table;

		$columns_table    = $this->columns_table;

		foreach($columns_table as $index => $coltab) {

			$field = @$coltab['name'];
			if(!$field) die('Please make sure there is key `name` in each row of col');

			if(strpos($field, ' as ')!==FALSE) {
				$field = substr($field, strpos($field, ' as ')+4);
				$result->addselect(DB::raw($coltab['name']));
				$columns_table[$index]['type_data']   = 'varchar';
				$columns_table[$index]['field']       = $field;
				$columns_table[$index]['field_raw']   = $field;
				$columns_table[$index]['field_with']  = $field;
				$columns_table[$index]['is_subquery'] = true;
				continue;
			}
			// dd(strpos($field),'.');

			if(strpos($field,'.')!==FALSE) {
				// dd('1');
				$result->addselect($field);
			}else{
				// dd('2');
				$result->addselect($table.'.'.$field);
			}

			$field_array = explode('.', $field);

			if(isset($field_array[1])) {
				$field = $field_array[1];
				$table = $field_array[0];
			}

			$columns_table[$index]['type_data']	 = CRUDBooster::getFieldType($table,$field);
			$columns_table[$index]['field']      = $field;
			$columns_table[$index]['field_raw']  = $field;
			$columns_table[$index]['field_with'] = $table.'.'.$field;

			$f = $this->findNameFormType($field);
			if ($f!==FALSE)
			{
				$columns_table[$index]['type_form'] = $f["type"];
				if($f["type"]=='select')
				{
					if (array_key_exists('datatable',$f))
					{
						$farr = explode(",",$f["datatable"]);
						$columns_table[$index]['optionlist'] = DB::table($farr[0])->pluck($farr[1])->toArray();							
					}
					else if (array_key_exists('dataenum',$f))
					{
						$farr = explode(";",$f["dataenum"]);
						$columns_table[$index]['optionlist'] = $farr;
					}
				}
			}

			
		}

		if(Request::get('q')) {
			$result->where(function($w) use ($columns_table, $request) {
				foreach($columns_table as $col) {
						if(!$col['field_with']) continue;
						if($col['is_subquery']) continue;
						$w->orwhere($col['field_with'],"like","%".Request::get("q")."%");
				}
			});
		}

		if(Request::get('where')) {
			foreach(Request::get('where') as $k=>$v) {
				$result->where($table.'.'.$k,$v);
			}
		}

		$filter_is_orderby = false;
		if(Request::get('filter_column')) {

			$filter_column = Request::get('filter_column');
			$result->where(function($w) use ($filter_column,$fc) {
				foreach($filter_column as $key=>$fc) {
					// dd($fc);
					$value = @$fc['value'];
					$type  = @$fc['type'];

					if($type == 'empty') {
						$w->whereNull($key)->orWhere($key,'');
						continue;
					}

					if($value == 'empty') {
						$w->whereNull($key)->orWhere($key,'');
					}

					if($value=='' || $type=='') continue;

					if($type == 'between') continue;

					switch($type) {
						default:
							if($key && $type && $value) $w->where($key,$type,$value);
						break;
						case 'like':
						case 'not like':
							$value = '%'.$value.'%';
							if($key && $type && $value) $w->where($key,$type,$value);
						break;
						case 'in':
						case 'not in':
							if($value) {
								$value = explode(',',$value);
								if($key && $value) $w->whereIn($key,$value);
							}
						break;
					}


				}
			});

			foreach($filter_column as $key=>$fc) {
				$value = @$fc['value'];
				$type  = @$fc['type'];
				$sorting = @$fc['sorting'];
				$type_data = @$fc['type_data'];
				$label_data = @$fc['label'];

				if($sorting!='') {
					if($key) {
						$result->orderby($key,$sorting);
						$filter_is_orderby = true;
					}
				}

				if ($type=='between') {
					
					if (($type_data == 'datetime')||((strpos(strtolower($key),"date")!==false)&&((strpos(strtolower($key),"time")!==false)||(strpos(strtolower($key),"create")!==false))))
					{	
							$value[0] .=" 00:00:00";
							$value[1] .=" 23:59:59";
							$result->whereBetween($key,$value);
					}
					else
						if($key && $value) $result->whereBetween($key,$value);
				}else{
					continue;
				}
			}
		}

		if($filter_is_orderby == true) {
			$data['result']  = $result->paginate($limit);

		}else{
			if($this->orderby) {
				if(is_array($this->orderby)) {
					foreach($this->orderby as $k=>$v) {
						if(strpos($k, '.')!==FALSE) {
							$orderby_table = explode(".",$k)[0];
						}else{
							$orderby_table = $table;
						}
						//$result->orderby($orderby_table.'.'.$k,$v);
						$result->orderby($k,$v);
					}
				}else{
					$this->orderby = explode(";",$this->orderby);
					foreach($this->orderby as $o) {
						$o = explode(",",$o);
						$k = $o[0];
						$v = $o[1];
						if(strpos($k, '.')!==FALSE) {
							$orderby_table = explode(".",$k)[0];
						}else{
							$orderby_table = $table;
						}
						//$result->orderby($orderby_table.'.'.$k,$v);
						$result->orderby($k,$v);
					}
				}
				$data['result'] = $result->paginate($limit);
			}else{
				$data['result'] = $result->orderby($this->table.'.'.$this->primary_key,'desc')->paginate($limit);
			}
		}

		echo "<pre>".print_r($data['result'])."</pre>";

		$data['columns'] = $columns_table;

		if($this->index_return) return $data;

		//LISTING INDEX HTML
		$addaction     = $this->data['addaction'];

		if($this->sub_module) {
			foreach($this->sub_module as $s) {
				$table_parent = CRUDBooster::parseSqlTable($this->table)['table'];
				$addaction[] = [
					'label'=>$s['label'],
					'icon'=>$s['button_icon'],
					'url'=>CRUDBooster::adminPath($s['path']).'?parent_table='.$table_parent.'&parent_columns='.$s['parent_columns'].'&parent_id=[id]&return_url='.urlencode(Request::fullUrl()).'&foreign_key='.$s['foreign_key'].'&label='.urlencode($s['label']),
					'color'=>$s['button_color'],
                                        'showIf'=>$s['showIf']
				];
			}
		}
		$
		$mainpath      = CRUDBooster::mainpath();
		$orig_mainpath = $this->data['mainpath'];
		$title_field   = $this->title_field;
		$html_contents = array();
		$page = (Request::get('page'))?Request::get('page'):1; 
		$number = ($page-1)*$limit+1; 
		foreach($data['result'] as $row) {
			$html_content = array();

			if($this->button_bulk_action) {		

				$html_content[] = "<input type='checkbox' class='checkbox' name='checkbox[]' value='".$row->{$tablePK}."'/>";
			}

			if($this->show_numbering) {
				$html_content[] = $number.'. ';
				$number++;
			}

			foreach($columns_table as $col) {
		          if($col['visible']===FALSE) continue;		          

		          $value = @$row->{$col['field']};
		          $title = @$row->{$this->title_field};
		          $label = $col['label'];

		          if(isset($col['image'])) {
			            if($value=='') {			              
			              $value = "<a  data-lightbox='roadtrip' rel='group_{{$table}}' title='$label: $title' href='http://placehold.it/50x50&text=NO+IMAGE'><img width='40px' height='40px' src='http://placehold.it/50x50&text=NO+IMAGE'/></a>";
			            }else{
							$pic = (strpos($value,'http://')!==FALSE)?$value:asset($value);				            
				            $value = "<a data-lightbox='roadtrip'  rel='group_{{$table}}' title='$label: $title' href='".$pic."'><img width='40px' height='40px' src='".$pic."'/></a>";
			            }			            
		          }

		          if(@$col['download']) {
			            $url = (strpos($value,'http://')!==FALSE)?$value:asset($value).'?download=1';
			            if($value) {
			            	$value = "<a class='btn btn-xs btn-primary' href='$url' target='_blank' title='Download File'><i class='fa fa-download'></i> Download</a>";
			            }else{
			            	$value = " - ";
			            }
		          }

		            if($col['str_limit']) {
		            	$value = trim(strip_tags($value));
		            	$value = str_limit($value,$col['str_limit']);
		            }

		            if($col['nl2br']) {
		            	$value = nl2br($value);
		            }

		            if($col['callback_php']) {
		              foreach($row as $k=>$v) {
		              		$col['callback_php'] = str_replace("[".$k."]",$v,$col['callback_php']);
		              }
		              @eval("\$value = ".$col['callback_php'].";");
		            }

		            //New method for callback
			        if(isset($col['callback'])) {
			        	$value = call_user_func($col['callback'],$row);
			        }


		            $datavalue = @unserialize($value);
					if ($datavalue !== false) {
						if($datavalue) {
							$prevalue = [];
							foreach($datavalue as $d) {
								if($d['label']) {
									$prevalue[] = $d['label'];
								}
						    }
						    if(count($prevalue)) {
						    	$value = implode(", ",$prevalue);
						    }
						}
					}

		          $html_content[] = $value;
	        } //end foreach columns_table


	      if($this->button_table_action):

	      		$button_action_style = $this->button_action_style;
	      		$html_content[] = "<div class='button_action' style='text-align:right'>".view('crudbooster::components.action',compact('addaction','row','button_action_style','parent_field'))->render()."</div>";

          endif;//button_table_action


          foreach($html_content as $i=>$v) {
          	$this->hook_row_index($i,$v);
          	$html_content[$i] = $v;
          }

	      $html_contents[] = $html_content;
		} //end foreach data[result]


		$customerlist = $this->getCustomer();

		// echo "<pre>".print_r($data,TRUE)."</pre><br>";

 		$html_contents = ['html'=>$html_contents,'data'=>$data['result']];

		$data['html_contents'] = $html_contents;
		//$data['limit'] = $result->count();
		//$data['limit'] = $totalCount;
		/*$itemSql = $result->toSql();
		$itemSql = str_replace("offset","start",$itemSql);
		echo $itemSql."<br>";*/


		return view("crudbooster::default.index",$data);
	
	} //last

	public function createTempTable()
	{
		$tableName = 'customergigya';
		
		$table = DB::insert(DB::raw("create table $tableName (
                                        id int NOT NULL AUTO_INCREMENT,
                                        firstName varchar(255) NOT NULL,
                                        lastName varchar(255) NOT NULL,
                                        email varchar(255) NOT NULL,
                                        PRIMARY KEY (id)
                                    )"));

		

		// $customerData = DB::table($tableName)->get();

		return $table;
	}


}