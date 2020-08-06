<?php namespace crocodicstudio\crudbooster\controllers;

error_reporting(E_ALL ^ E_NOTICE ^ E_WARNING);


use crocodicstudio\crudbooster\controllers\Controller;
use App\Customer;
use App\Mainmerge;
use App\DbtWhatsappNumber;
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
use DateTime;
use Carbon;
use App\Traits\GigyaApi;
use App\Traits\SetSmartDataInfoToGigya;
use Config;
use GuzzleHttp\Exception\GuzzleException;
use GuzzleHttp\Client;
use App\Http\Controllers\Alerts\SMSAlertController;

class CBController extends Controller {

	use GigyaApi;
	use SetSmartDataInfoToGigya;

	public $data_inputan;
	public $columns_table;
	public $module_name;
	public $table;
	public $title_field;
	public $primary_key        = 'id';
	public $arr                = array();
	public $col                = array();
	public $form               = array();
	public $data               = array();
	public $addaction          = array();
	public $orderby            = NULL;
	public $password_candidate = NULL;
	public $date_candidate     = NULL;
	public $limit              = 20;
	public $global_privilege   = FALSE;
	public $show_numbering	   = FALSE;

	public $alert                 = array();
	public $index_button          = array();

	public $button_filter       = TRUE;
	public $button_export       = TRUE;
	public $button_import       = FALSE;
	public $button_show         = TRUE;
	public $button_addmore      = TRUE;
	public $button_table_action = TRUE;
	public $button_bulk_action	= TRUE;
	public $button_add          = TRUE;
	public $button_delete       = TRUE;
	public $button_cancel       = TRUE;
	public $button_save         = TRUE;
	public $button_edit         = TRUE;
	public $button_detail       = TRUE;
	public $button_action_style = 'button_icon';
	public $button_action_width = NULL;

	public $index_statistic       = array();
	public $index_additional_view = array();
	public $pre_index_html        = null;
	public $post_index_html       = null;
	public $load_js               = array();
	public $load_css              = array();
	public $script_js             = NULL;
	public $style_css             = NULL;
	public $sub_module            = array();
	public $show_addaction        = TRUE;
	public $table_row_color 	  = array();
	public $button_selected 	  = array();
	public $return_url 			  = NULL;
	public $parent_field 		  = NULL;
	public $parent_id 			  = NULL;
	public $hide_form			  = array();
	public $index_return 		  = FALSE; //for export
	public $option_id			  = FALSE;
	public $option_fields		  = array();
	public $import_consignment	  = FALSE;
	public $gigya_based			  = FALSE;
	public $gigya_customer        = FALSE;
	public $thailand_customer     = FALSE;
	public $import_mobile_number  = FALSE;
	public $import_offline	      = FALSE;
	public $sfmc_alert            = FALSE;

	public function __construct()
	{
		$this->gigya_api_key    = config('gigyaaccess.GIGYAAPIKEY');
		$this->gigya_secret_key = config('gigyaaccess.GIGYASECRETKEY');
		$this->gigya_user_key   = config('gigyaaccess.GIGYAUSERKEY');
	}

	public function cbLoader() {
		$this->cbInit();

		$this->checkHideForm();

		$this->primary_key 					 = CB::pk($this->table);
		$this->columns_table                 = $this->col;
		$this->data_inputan                  = $this->form;
		$this->data['pk']				     = $this->primary_key;
		$this->data['forms']                 = $this->data_inputan;
		$this->data['hide_form'] 			 = $this->hide_form;
		$this->data['addaction']             = ($this->show_addaction)?$this->addaction:NULL;
		$this->data['table']                 = $this->table;
		$this->data['title_field']           = $this->title_field;
		$this->data['appname']               = CRUDBooster::getSetting('appname');
		$this->data['alerts']                = $this->alert;
		$this->data['index_button']          = $this->index_button;
		$this->data['show_numbering']	     = $this->show_numbering;
		$this->data['button_detail']         = $this->button_detail;
		$this->data['button_edit']           = $this->button_edit;
		$this->data['button_show']           = $this->button_show;
		$this->data['button_add']            = $this->button_add;
		$this->data['button_delete']         = $this->button_delete;
		$this->data['button_filter']         = $this->button_filter;
		$this->data['button_export']         = $this->button_export;
		$this->data['button_addmore']        = $this->button_addmore;
		$this->data['button_cancel']         = $this->button_cancel;
		$this->data['button_save']           = $this->button_save;
		$this->data['button_table_action']   = $this->button_table_action;
		$this->data['button_bulk_action']    = $this->button_bulk_action;
		$this->data['button_import']         = $this->button_import;
		$this->data['button_action_width']   = $this->button_action_width;
		$this->data['button_selected']       = $this->button_selected;
		$this->data['index_statistic']       = $this->index_statistic;
		$this->data['index_additional_view'] = $this->index_additional_view;
		$this->data['table_row_color']       = $this->table_row_color;
		$this->data['pre_index_html']        = $this->pre_index_html;
		$this->data['post_index_html']       = $this->post_index_html;
		$this->data['load_js']               = $this->load_js;
		$this->data['load_css']              = $this->load_css;
		$this->data['script_js']             = $this->script_js;
		$this->data['style_css']             = $this->style_css;
		$this->data['sub_module']            = $this->sub_module;
		$this->data['parent_field'] 		 = (g('parent_field'))?:$this->parent_field;
		$this->data['parent_id'] 		 	 = (g('parent_id'))?:$this->parent_id;

		if(CRUDBooster::getCurrentMethod() == 'getProfile') {
			Session::put('current_row_id',CRUDBooster::myId());
			$this->data['return_url'] = Request::fullUrl();
		}

        view()->share($this->data);
	}

	public function cbView($template,$data) {
		$this->cbLoader();
		echo view($template,$data);
	}

	protected function checkHideForm() {
		if(count($this->hide_form)) {
			foreach($this->form as $i=>$f) {
				if(in_array($f['name'], $this->hide_form)) {
					unset($this->form[$i]);
				}
			}
		}
	}

	protected function findNameFormType($name) {
		$isFind = FALSE;
		foreach($this->form as $i=>$f)
		{
			if ($f["name"]==$name)
				$isFind = $f;
		}
		return $isFind;
	}

	public function checkAccess() {
		$module = CRUDBooster::getCurrentModule();

		if(!CRUDBooster::isView() && $this->global_privilege==FALSE) {
			CRUDBooster::insertLog(trans('crudbooster.log_try_view',['module'=>$module->name]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans('crudbooster.denied_access'));
		}
	}

	public function processParents()
	{
		$data = [];
		if(Request::get('parent_table')) {
			$parentTablePK = CB::pk(g('parent_table'));
			$data['parent_table'] = DB::table(Request::get('parent_table'))->where($parentTablePK,Request::get('parent_id'))->first();
			if(Request::get('foreign_key')) {
				$data['parent_field'] = Request::get('foreign_key');
			}else{
				$data['parent_field'] = CB::getTableForeignKey(g('parent_table'),$this->table);
			}

			if($parent_field) {
				foreach($this->columns_table as $i=>$col) {
					if($col['name'] == $parent_field) {
						unset($this->columns_table[$i]);
					}
				}
			}			
		}
		return $data;
	}

	public function generateQuery()
	{
		$result = DB::table($this->table)->select(DB::raw($this->table.".".$this->primary_key));
		if(Request::get('parent_id')) {
			$table_parent = $this->table;
			$table_parent = CRUDBooster::parseSqlTable($table_parent)['table'];
			$result->where($table_parent.'.'.Request::get('foreign_key'),Request::get('parent_id'));
		}
		return $result;
	}

	public function checkParentsColumnsAndMerge($columns_table)
	{
		if(Request::get('parent_columns'))
		{
			$parentColumns = explode(",",Request::get('parent_columns'));
			foreach($columns_table as $key => $column)
			{
				if(in_array($column['name'], $parentColumns))
	            {
	              unset($columns_table[$key]);
	            }
			}
		}

		$columns_table = array_merge($columns_table);
		return $columns_table;
	}

	public function generateSubqueryField($coltab,&$columns_table, &$result,$index)
	{
		$field = substr($field, strpos($field, ' as ')+4);
		$result->addselect(DB::raw($coltab['name']));
		$columns_table[$index]['type_data']   = 'varchar';
		$columns_table[$index]['field']       = $field;
		$columns_table[$index]['field_raw']   = $field;
		$columns_table[$index]['field_with']  = $field;
		$columns_table[$index]['is_subquery'] = true;		
	}

	public function processColumnsQuery(&$columns_table, &$result)
	{
		$table            = $this->table;
		$join_alias_count = 0;
		foreach($columns_table as $index => $coltab) {
			$join = @$coltab['join'];
			$join_where = @$coltab['join_where'];
			$join_id = @$coltab['join_id'];
			$field = @$coltab['name'];
			$join_table_temp[] = $table;

			if(!$field) die('Please make sure there is key `name` in each row of col');

			if(strpos($field, ' as ')!==FALSE) {				
				$this->generateSubqueryField($coltab, $columns_table,$result,$index);								
				continue;
			}

			else if(strpos($field,'.')!==FALSE) {
				$result->addselect($field);
			}else{				
				$result->addselect($table.'.'.$field);
			}

			$field_array = explode('.', $field);

			if(isset($field_array[1])) {
				$field = $field_array[1];
				$table = $field_array[0];
			}

			if($join) {

				$join_exp    = explode(',', $join);

				$join_table  = $join_exp[0];
				$joinTablePK = CB::pk($join_table);
				$join_column = $join_exp[1];
				$join_alias  = str_replace(".", "_", $join_table);

				if(in_array($join_table, $join_table_temp)) {
					$join_alias_count += 1;
					$join_alias = $join_table.$join_alias_count;
				}
				$join_table_temp[] = $join_table;

				$result->leftjoin($join_table.' as '.$join_alias,$join_alias.(($join_id)? '.'.$join_id:'.'.$joinTablePK),'=',DB::raw($table.'.'.$field. (($join_where) ? ' AND '.$join_where.' ':'') ) );
				$result->addselect($join_alias.'.'.$join_column.' as '.$join_alias.'_'.$join_column);

				$join_table_columns = CRUDBooster::getTableColumns($join_table);
				if($join_table_columns) {
					foreach($join_table_columns as $jtc) {
						$result->addselect($join_alias.'.'.$jtc.' as '.$join_alias.'_'.$jtc);
					}
				}

				$alias[] = $join_alias;
				$columns_table[$index]['type_data']	 = CRUDBooster::getFieldType($join_table,$join_column);
				$columns_table[$index]['field']      = $join_alias.'_'.$join_column;
				$columns_table[$index]['field_with'] = $join_alias.'.'.$join_column;
				$columns_table[$index]['field_raw']  = $join_column;

				@$join_table1  = $join_exp[2];
				@$joinTable1PK = CB::pk($join_table1);
				@$join_column1 = $join_exp[3];
				@$join_alias1  = $join_table1;

				if($join_table1 && $join_column1) {

					if(in_array($join_table1, $join_table_temp)) {
						$join_alias_count += 1;
						$join_alias1 = $join_table1.$join_alias_count;
					}

					$join_table_temp[] = $join_table1;

					$result->leftjoin($join_table1.' as '.$join_alias1,$join_alias1.'.'.$joinTable1PK,'=',$join_alias.'.'.$join_column);
					$result->addselect($join_alias1.'.'.$join_column1.' as '.$join_column1.'_'.$join_alias1);
					$alias[] = $join_alias1;
					$columns_table[$index]['type_data']	 = CRUDBooster::getFieldType($join_table1,$join_column1);
					$columns_table[$index]['field']      = $join_column1.'_'.$join_alias1;
					$columns_table[$index]['field_with'] = $join_alias1.'.'.$join_column1;
					$columns_table[$index]['field_raw']  = $join_column1;
				}

			}else{

				//$result->addselect($table.'.'.$field);
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
		}
	}

	public function processFilterColumns(&$columns_table, &$result,&$data)
	{
		$table            = $this->table;
		$limit			  = $data["limit"];		

		$filter_is_orderby = false;
		if(Request::get('filter_column')) {			
			$filter_column = Request::get('filter_column');
			$result->where(function($w) use ($columns_table,$filter_column,$fc) {
				foreach($filter_column as $key=>$fc) {

					if($key) {
						$icheck = false;
						foreach($columns_table as $c)
						{
							if ($c['field']==$key)
							{
								if (strpos($c['name'], " as ") !== false) {
										$key = "(".substr($c['name'], 0, strpos($c['name'], " as ")).")";
										$icheck = true;										
										break;
								}
							}
						}
					}
					
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
							if (!$icheck)
								if($key && $type && $value) $w->where($key,$type,$value);
							else
								if($key && $type && $value) $w->whereRaw($key." ".$type." ".$value);
						break;
						case 'in':
						case 'not in':
							if($value) {
								$value2 = explode(',',$value);
								if (!$icheck)
									if($key && $value2) $w->whereIn($key,$value2);
								else
								{
									if($key && $value2) $w->whereRaw($key." IN (".$value.")");
								}
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

				if($key) {
					$icheck = false;
					foreach($columns_table as $c)
					{
						if ($c['field']==$key)
						{
							if (strpos($c['name'], " as ") !== false) {
									$key = "(".substr($c['name'], 0, strpos($c['name'], " as ")).")";
									$icheck = true;
									break;
							}
						}
					}
				}


				if($sorting!='') {
					if($key) {


						if (!$icheck)
							$result->orderby($key,$sorting);
						else {
							$result->orderByRaw($key." ".$sorting);
						}
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
						$result->orderby($orderby_table.'.'.$k,$v);
					}
				}else{
					//dd($this->orderby);
					$this->orderby = explode(";",$this->orderby);
					foreach($this->orderby as $o) {
						$o = explode(",",$o);
						$k = $o[0];
						$v = $o[1];
						if(strpos($k, '.')!==FALSE) {
							$orderby_table = explode(".",$k)[0];
							$orderby_field = explode(".",$k)[1];							
							if ($orderby_table=="")
								$orderby_table = $this->table;							
							$result->orderby($orderby_table.'.'.$orderby_field,$v);
						}else{
							$orderby_table = $this->table;
							$result->orderby($orderby_table.'.'.$k,$v);
						}

					}
				}
				//dd($result);

				$data['result'] = $result->paginate($limit);
			}else{

				$data['result'] = $result->orderby($this->table.'.'.$this->primary_key,'desc')->paginate($limit);

			}
		}

		$data['columns'] = $columns_table;
	}

	public function initDataValues()
	{
		$data['table'] 	  = $this->table;
		$data['table_pk'] = CB::pk($this->table);
		$data['page_title']       = $module->name;
		$data['page_description'] = trans('crudbooster.default_module_description');
		$data['date_candidate']   = $this->date_candidate;
		$data['limit'] = $limit   = (Request::get('limit'))?Request::get('limit'):$this->limit;
		return $data;
	}

	public function processSearchQuery(&$columns_table,&$result)
	{
		if(Request::get('q')) {
			$result->where(function($w) use ($columns_table, $request) {
				foreach($columns_table as $col) {
						if(!$col['field_with']) continue;
						if($col['is_subquery']) continue;
						$w->orwhere($col['field_with'],"like","%".Request::get("q")."%");
				}
			});
		}
	}

	public function processWhereQuery(&$columns_table,&$result)
	{
		$table 	  = $this->table;
		if(Request::get('where')) {
			foreach(Request::get('where') as $k=>$v) {
				$result->where($table.'.'.$k,$v);
			}
		}
	}

	public function processCustomActions()
	{
		$addaction     = $this->data['addaction'];

		if($this->sub_module) {
			foreach($this->sub_module as $s) {
				$table_parent = CRUDBooster::parseSqlTable($this->table)['table'];
				$addaction[] = [
					'label'=>$s['label'],
					'icon'=>$s['button_icon'],
					'url'=>CRUDBooster::adminPath($s['path']).'?parent_table='.$table_parent.'&parent_columns='.$s['parent_columns'].'&custom_parent_alias='.$s['custom_parent_alias'].'&parent_id=[id]&return_url='.urlencode(Request::fullUrl()).'&foreign_key='.$s['foreign_key'].'&label='.urlencode($s['label']),
					'color'=>$s['button_color'],
                                        'showIf'=>$s['showIf']
				];
			}
		}

		return $addaction;
		
	}

	public function getIndex() {
		$this->cbLoader();
		$this->checkAccess();		

		$data = $this->processParents();
		$data = $this->initDataValues();
	
		$tablePK = $data['table_pk'];
		$table_columns = CB::getTableColumns($this->table);
		$result  = $this->generateQuery();	
		
		//dd($result);

		$this->hook_query_index($result);

		if(in_array('deleted_at', $table_columns)) {
			$result->where($this->table.'.deleted_at',NULL);
		}

		$alias            = array();
		$join_alias_count = 0;
		$join_table_temp  = array();
		$table            = $this->table;
		$columns_table    = $this->columns_table;

		$columns_table = $this->checkParentsColumnsAndMerge($columns_table);				
		$this->processColumnsQuery($columns_table,$result);		
		$this->processSearchQuery($columns_table,$result);
		$this->processWhereQuery($columns_table,$result);				
		$this->processFilterColumns($columns_table,$result,$data);		

		if($this->index_return) return $data;

		//LISTING INDEX HTML
		$addaction = $this->processCustomActions();		

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
	      		$onlyView            = false;
	      		$html_content[]      = "<div class='button_action' style='text-align:right'>".view('crudbooster::components.action',compact('addaction','row','button_action_style','parent_field', 'onlyView'))->render()."</div>";

          endif;//button_table_action

          foreach($html_content as $i=>$v) {
          	$this->hook_row_index($i,$v);
          	$html_content[$i] = $v;
          }

	      $html_contents[] = $html_content;
		} //end foreach data[result]

 		$html_contents = ['html'=>$html_contents,'data'=>$data['result']];

		$data['html_contents'] = $html_contents;
		$data['limit'] = $result->count();
		//echo $result->toSql()."<br>";
		return view("crudbooster::default.index",$data);
	}

	public function getExportData() {

		return redirect(CRUDBooster::mainpath());
	}

	public function postExportData() {
		$this->limit 		= Request::input('limit');
		$this->index_return = TRUE;
		$filetype 			= Request::input('fileformat');
		$filename 			= Request::input('filename');
		$papersize			= Request::input('page_size');
		$paperorientation	= Request::input('page_orientation');
		$response           = $this->getIndex();

		if(Request::input('default_paper_size')) {
			DB::table('cms_settings')->where('name','default_paper_size')->update(['content'=>$papersize]);
		}

		$userName = CRUDBooster::myName();

		$description = $userName.' has export data from '.CRUDBooster::getCurrentModule()->name;

		CRUDBooster::insertLog($description);

		switch($filetype) {
			case "pdf":
				$view = view('crudbooster::export',$response)->render();
				$pdf = App::make('dompdf.wrapper');
				$pdf->loadHTML($view);
				$pdf->setPaper($papersize,$paperorientation);
				return $pdf->stream($filename.'.pdf');
			break;
			case 'xls':
				Excel::create($filename, function($excel) use ($response) {
					$excel->setTitle($filename)
					->setCreator("crudbooster.com")
					->setCompany(CRUDBooster::getSetting('appname'));
				    $excel->sheet($filename, function($sheet) use ($response) {
				    	$sheet->setOrientation($paperorientation);
				        $sheet->loadview('crudbooster::export',$response);
				    });
				})->export('xls');
			break;
			case 'csv':
				Excel::create($filename, function($excel) use ($response) {
					$excel->setTitle($filename)
					->setCreator("crudbooster.com")
					->setCompany(CRUDBooster::getSetting('appname'));
				    $excel->sheet($filename, function($sheet) use ($response) {
				    	$sheet->setOrientation($paperorientation);
				        $sheet->loadview('crudbooster::export',$response);
				    });
				})->export('csv');
			break;
		}
	}

	public function postDataQuery() {
		$query = Request::get('query');
		$query = DB::select(DB::raw($query));
		return response()->json($query);
	}

	public function getDataTable() {
		$table = Request::get('table');
		$label = Request::get('label');
		$foreign_key_name = Request::get('fk_name');
		$foreign_key_value = Request::get('fk_value');
		if($table && $label && $foreign_key_name && $foreign_key_value) {
			$query = DB::table($table)->select('id as select_value',$label.' as select_label')->where($foreign_key_name,$foreign_key_value)->orderby($label,'asc')->get();

			//$query = DB::table($table)->select($label.' as select_value',$label.' as select_label')->where($foreign_key_name,$foreign_key_value)->orderby($label,'asc')->get();
			return response()->json($query);
		}else{
			return response()->json([]);
		}
	}

	public function getModalData() {
		$table = Request::get('table');
		$where = Request::get('where');
		$where = urldecode($where);
		$columns = Request::get('columns');
		$columns = explode(",",$columns);

		$table = CRUDBooster::parseSqlTable($table)['table'];
		$tablePK = CB::pk($table);
		$result = DB::table($table);

		if(Request::get('q')) {
			$result->where(function($where) use ($columns) {
				foreach($columns as $c=>$col) {
					if($c==0) {
						$where->where($col,'like','%'.Request::get('q').'%');
					}else{
						$where->orWhere($col,'like','%'.Request::get('q').'%');
					}
				}
			});
		}

		if($where) {
			$result->whereraw($where);
		}

		$result->orderby($tablePK,'desc');

		$data['result'] = $result->paginate(6);
		$data['columns'] = $columns;
		return view('crudbooster::default.type_components.datamodal.browser',$data);
	}

	public function getUpdateSingle() {
		$table = Request::get('table');
		$column = Request::get('column');
		$value = Request::get('value');
		$id = Request::get('id');
		$tablePK = CB::pk($table);
		DB::table($table)->where($tablePK,$id)->update([$column => $value]);

		return redirect()->back()->with(['message_type'=>'success','message'=>trans('crudbooster.alert_delete_data_success')]);
	}

	public function getFindData() {
		$q        = Request::get('q');
		$id       = Request::get('id');
		$limit    = Request::get('limit')?:10;
		$format   = Request::get('format');

		$table1   = (Request::get('table1'))?:$this->table;
		$table1PK = CB::pk($table1);
		$column1  = (Request::get('column1'))?:$this->title_field;

		@$table2  = Request::get('table2');
		@$column2 = Request::get('column2');

		@$table3  = Request::get('table3');
		@$column3 = Request::get('column3');

		$where    = Request::get('where');

		$fk 	  = Request::get('fk');
		$fk_value = Request::get('fk_value');

		if($q || $id || $table1) {
			$rows = DB::table($table1);
			$rows->select($table1.'.*');
			$rows->take($limit);

			if(CRUDBooster::isColumnExists($table1,'deleted_at')) {
				$rows->where($table1.'.deleted_at',NULL);
			}

			if($fk && $fk_value) {
				$rows->where($table1.'.'.$fk,$fk_value);
			}

			if($table1 && $column1) {

				$orderby_table  = $table1;
				$orderby_column = $column1;
			}

			if($table2 && $column2) {
				$table2PK = CB::pk($table2);
				$rows->join($table2,$table2.'.'.$table2PK,'=',$table1.'.'.$column1);
				$columns = CRUDBooster::getTableColumns($table2);
				foreach($columns as $col) {
					$rows->addselect($table2.".".$col." as ".$table2."_".$col);
				}
				$orderby_table  = $table2;
				$orderby_column = $column2;
			}

			if($table3 && $column3) {
				$table3PK = CB::pk($table3);
				$rows->join($table3,$table3.'.'.$table3PK,'=',$table2.'.'.$column2);
				$columns = CRUDBooster::getTableColumns($table3);
				foreach($columns as $col) {
					$rows->addselect($table3.".".$col." as ".$table3."_".$col);
				}
				$orderby_table  = $table3;
				$orderby_column = $column3;
			}

			if($id) {
				$rows->where($table1.".".$table1PK,$id);
			}

			if($where) {
				$rows->whereraw($where);
			}

			if($format) {
				$format = str_replace('&#039;', "'", $format);
				$rows->addselect(DB::raw("CONCAT($format) as text"));
				if($q) $rows->whereraw("CONCAT($format) like '%".$q."%'");
			}else{
				$rows->addselect($orderby_table.'.'.$orderby_column.' as text');
				if($q) $rows->where($orderby_table.'.'.$orderby_column,'like','%'.$q.'%');
			}

			$result          = array();
			$result['items'] = $rows->get();
		}else{
			$result          = array();
			$result['items'] = array();
		}
		return response()->json($result);
	}

	public function validation($id=NULL, $isjs=false) {


		$request_all = Request::all();
		$array_input = array();


		foreach($this->data_inputan as $di) {
			$ai = array();
			$name = $di['name'];

			if (( !isset($request_all[$name]) ) && ($isjs==false)) continue;

			if($di['type'] != 'upload') {
				if(@$di['required']) {
					$ai[] = 'required';
				}
			}

			if($di['type'] == 'upload') {
				if($id) {
					$row = DB::table($this->table)->where($this->primary_key,$id)->first();
					if($row->{$di['name']}=='') {
						$ai[] = 'required';
					}
				}
			}

			if(@$di['min']) {
				$ai[] = 'min:'.$di['min'];
			}
			if(@$di['max']) {
				$ai[] = 'max:'.$di['max'];
			}
			if(@$di['image']) {
				$ai[] = 'image';
			}
			if(@$di['mimes']) {
				$ai[] = 'mimes:'.$di['mimes'];
			}
			$name = $di['name'];
			if(!$name) continue;

			if($di['type']=='money') {
				$request_all[$name] = preg_replace('/[^\d-]+/', '', $request_all[$name]);
			}


			if(@$di['validation']) {

				$exp = explode('|',$di['validation']);
				if (count($exp)) {
					foreach ($exp as &$validationItem) {
						if (substr($validationItem, 0,6) == 'unique') {
							$parseUnique = explode(',',str_replace('unique:','',$validationItem));
							$uniqueTable = ($parseUnique[0])?:$this->table;
							$uniqueColumn = ($parseUnique[1])?:$name;
							$uniqueIgnoreId = ($parseUnique[2])?:(($id)?:'');

							//Make sure table name
							$uniqueTable = CB::parseSqlTable($uniqueTable)['table'];

							//Rebuild unique rule
							$uniqueRebuild = [];
							$uniqueRebuild[] = $uniqueTable;
							$uniqueRebuild[] = $uniqueColumn;
							if ($uniqueIgnoreId) {
								$uniqueRebuild[] = $uniqueIgnoreId;
							} else {
								$uniqueRebuild[] = 'NULL';
							}

							//Check whether deleted_at exists or not
							if (CB::isColumnExists($uniqueTable,'deleted_at')) {
								$uniqueRebuild[] = CB::findPrimaryKey($uniqueTable);
								$uniqueRebuild[] = 'deleted_at';
								$uniqueRebuild[] = 'NULL';
							}
							$uniqueRebuild = array_filter($uniqueRebuild);
							$validationItem = 'unique:'.implode(',',$uniqueRebuild);
						}
						if (substr($validationItem, 0,9) == 'sixmonths') {
							$parseSixMonths = explode(',',str_replace('sixmonths:','',$validationItem));
							if (isset($id))
							{
								$parseSixMonths[] = $id;
							}
							$validationItem = 'sixmonths:'.implode(',',$parseSixMonths);
						}
					}
				} else {
					$exp = array();
				}

				$validation = implode('|',$exp);

				$array_input[$name] = $validation;
			}else{
				$array_input[$name] = implode('|',$ai);
			}
		}

		if ($isjs)
		{
			$validator = JsValidator::make($array_input);

			return $validator;
		}
		else
		{
			$validator = Validator::make($request_all,$array_input);

			if ($validator->fails())
			{
				$message = $validator->messages();
				$message_all = $message->all();

				if(Request::ajax()) {
					$res = response()->json(['message'=>trans('crudbooster.alert_validation_error',['error'=>implode(', ',$message_all)]),'message_type'=>'warning'])->send();
					exit;
				}else{
					$res = redirect()->back()->with("errors",$message)->with(['message'=>trans('crudbooster.alert_validation_error',['error'=>implode(', ',$message_all)]),'message_type'=>'warning'])->withInput();
					\Session::driver()->save();
					$res->send();
		        	exit;
				}

			}
		}
	}

	public function validationArray($formarray, $isjs=false) {

		$request_all = $formarray;
		$array_input = array();
		foreach($this->form as $di) {
			$ai = array();
			$name = $di['name'];

			if( !isset($formarray[$name]) ) continue;

			if($di['type'] != 'upload') {
				if(@$di['required']) {
					$ai[] = 'required';
				}
			}

			if($di['type'] == 'upload') {
				if($id) {
					$row = DB::table($this->table)->where($this->primary_key,$id)->first();
					if($row->{$di['name']}=='') {
						$ai[] = 'required';
					}
				}
			}

			if(@$di['min']) {
				$ai[] = 'min:'.$di['min'];
			}
			if(@$di['max']) {
				$ai[] = 'max:'.$di['max'];
			}
			if(@$di['image']) {
				$ai[] = 'image';
			}
			if(@$di['mimes']) {
				$ai[] = 'mimes:'.$di['mimes'];
			}
			$name = $di['name'];
			if(!$name) continue;

			if($di['type']=='money') {
				$formarray[$name] = preg_replace('/[^\d-]+/', '', $formarray[$name]);
			}


			if(@$di['validation']) {

				$exp = explode('|',$di['validation']);
				if (count($exp)) {
					foreach ($exp as &$validationItem) {
						if (substr($validationItem, 0,6) == 'unique') {
							$parseUnique = explode(',',str_replace('unique:','',$validationItem));
							$uniqueTable = ($parseUnique[0])?:$this->table;
							$uniqueColumn = ($parseUnique[1])?:$name;
							$uniqueIgnoreId = ($parseUnique[2])?:(($id)?:'');

							//Make sure table name
							$uniqueTable = CB::parseSqlTable($uniqueTable)['table'];

							//Rebuild unique rule
							$uniqueRebuild = [];
							$uniqueRebuild[] = $uniqueTable;
							$uniqueRebuild[] = $uniqueColumn;
							if ($uniqueIgnoreId) {
								$uniqueRebuild[] = $uniqueIgnoreId;
							} else {
								$uniqueRebuild[] = 'NULL';
							}

							//Check whether deleted_at exists or not
							if (CB::isColumnExists($uniqueTable,'deleted_at')) {
								$uniqueRebuild[] = CB::findPrimaryKey($uniqueTable);
								$uniqueRebuild[] = 'deleted_at';
								$uniqueRebuild[] = 'NULL';
							}
							$uniqueRebuild = array_filter($uniqueRebuild);
							$validationItem = 'unique:'.implode(',',$uniqueRebuild);
						}
					}
				} else {
					$exp = array();
				}


				$validation = implode('|',$exp);

				$array_input[$name] = $validation;
			}else{
				$array_input[$name] = implode('|',$ai);
			}
		}

		if (!$isjs)
			$validator = Validator::make($formarray,$array_input);
		else
			$validator = JsValidator::make($formarray,$array_input);

		return $validator;
	}

	public function input_assignment($id=null) {

		$hide_form = (Request::get('hide_form'))?unserialize(Request::get('hide_form')):array();

		foreach($this->data_inputan as $ro) {
			$name = $ro['name'];
			if(!$name) continue;

			if($ro['exception']) continue;

			if($name=='hide_form') continue;

			if(count($hide_form)) {
				if(in_array($name, $hide_form)) {
					continue;
				}
			}

			if($ro['type']=='checkbox' && $ro['relationship_table']) {
				continue;
			}

			if($ro['type']=='select2' && $ro['relationship_table']) {
				continue;
			}

			$inputdata = Request::get($name);

			if($ro['type']=='money') {
				$inputdata = preg_replace('/[^\d-]+/', '', $inputdata);
			}

			if($ro['type']=='child') continue;

			if($name) {
				if($inputdata!='') {
					$this->arr[$name] = $inputdata;
				}else{
					$this->arr[$name] = "";
				}
			}

			$password_candidate = explode(',',config('crudbooster.PASSWORD_FIELDS_CANDIDATE'));
			if(in_array($name, $password_candidate)) {
				if(!empty($this->arr[$name])) {
					$this->arr[$name] = Hash::make($this->arr[$name]);
				}else{
					unset($this->arr[$name]);
				}
			}

			if($ro['type']=='checkbox') {

				if(is_array($inputdata)) {
					if($ro['datatable'] != '') {
						$table_checkbox = explode(',',$ro['datatable'])[0];
						$field_checkbox = explode(',',$ro['datatable'])[1];
						$table_checkbox_pk = CB::pk($table_checkbox);
						$data_checkbox = DB::table($table_checkbox)->whereIn($table_checkbox_pk,$inputdata)->pluck($field_checkbox)->toArray();
						$this->arr[$name] = implode(";",$data_checkbox);
					}else{
						$this->arr[$name] = implode(";",$inputdata);
					}
				}
			}

			//multitext colomn
			if($ro['type']=='multitext') {
				$name = str_slug($ro['name'],'');
				$multitext="";

				for($i=0;$i<=count($this->arr[$name])-1;$i++) {
					$multitext .= $this->arr[$name][$i]."|";
				}
				$multitext=substr($multitext,0,strlen($multitext)-1);
				$this->arr[$name]=$multitext;
			}

			if($ro['type']=='googlemaps') {
				if($ro['latitude'] && $ro['longitude']) {
					$latitude_name = $ro['latitude'];
					$longitude_name = $ro['longitude'];
					$this->arr[$latitude_name] = Request::get('input-latitude-'.$name);
					$this->arr[$longitude_name] = Request::get('input-longitude-'.$name);
				}
			}

			if($ro['type']=='select' || $ro['type']=='select2') {
				if($ro['datatable']) {
					if(($inputdata=='') && (($this->option_id === FALSE)||(in_array($name, $this->option_fields)))) {
						$this->arr[$name] = 0;
					}
				}
			}


			if(@$ro['type']=='upload') {
				if (Request::hasFile($name))
				{
					$file = Request::file($name);
					$ext  = $file->getClientOriginalExtension();
					$filename = str_slug(pathinfo($file->getClientOriginalName(), PATHINFO_FILENAME));

					//Create Directory Monthly
					Storage::makeDirectory(date('Y-m'));

					//Move file to storage
					$file_path = storage_path('app'.DIRECTORY_SEPARATOR.date('Y-m'));

					if($ro['upload_encrypt']==true) {
						$filename = md5(str_random(5)).'.'.$ext;
					}else{
						if(count(glob($file_path.'/'.$filename))>0)
						{
							$filename = $filename.'_'.count(glob($file_path."/$filename*.$ext")).'.'.$ext;
						}else{
							$filename = $filename.'.'.$ext;
						}
					}

					if($file->move($file_path,$filename)) {
						$this->arr[$name] = 'uploads/'.date('Y-m').'/'.$filename;
					}
				}

				if(!$this->arr[$name]) {
					$this->arr[$name] = Request::get('_'.$name);
				}
			}

			if(@$ro['type']=='filemanager') {
				$url = str_replace(asset('/'),'',$this->arr[$name]);
				$url = str_replace("//","/",$url);
				$this->arr[$name] = $url;
			}
		}
	}

	public function getAdd(){
		$this->cbLoader();
		if(!CRUDBooster::isCreate() && $this->global_privilege==FALSE || $this->button_add==FALSE) {
			CRUDBooster::insertLog(trans('crudbooster.log_try_add',['module'=>CRUDBooster::getCurrentModule()->name ]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans("crudbooster.denied_access"));
		}

		$page_title      = trans("crudbooster.add_data_page_title",['module'=>CRUDBooster::getCurrentModule()->name]);
		$page_menu       = Route::getCurrentRoute()->getActionName();
		$command 		 = 'add';
		$option_id		 = $this->option_id;
		$option_fields	 = $this->option_fields;
		$validator		 = $this->validation(NULL,true);
		$gigya_customer  = $this->gigya_customer;

		return view('crudbooster::default.form',compact('page_title','page_menu','command','option_id','validator','option_fields', 'gigya_customer'));
	}

	public function postAddSave() {
		$this->cbLoader();

		if(!CRUDBooster::isCreate() && $this->global_privilege==FALSE) {
			CRUDBooster::insertLog(trans('crudbooster.log_try_add_save',['name'=>Request::input($this->title_field),'module'=>CRUDBooster::getCurrentModule()->name ]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans("crudbooster.denied_access"));
		}

		$this->validation();
		$this->input_assignment();

		if(Schema::hasColumn($this->table, 'created_at'))
		{
		    $this->arr['created_at'] = date('Y-m-d H:i:s');
		}

		if($this->table == 'cms_users')
		{
			$this->arr['password_updated_at'] = date('Y-m-d H:i:s');
		}

		$this->hook_before_add($this->arr);

		$this->arr[$this->primary_key] = $id = CRUDBooster::newId($this->table);

		if($this->gigya_customer)
		{
			$this->arr['is_gigya_customer'] = 1;
		}

		if(isset($this->arr['subsource_id']))
		{
			$this->arr['m_subsource'] = $this->arr['subsource_id'];
			unset($this->arr['subsource_id']);
		}

		if($this->gigya_customer)
		{
			$recordId = DB::table($this->table)->insertGetId($this->arr);

			if($this->arr['countryCode'] = 'SG')
			{
				$subsampleRecords = ['m_product' => 'First Welcome Gift', 'stock_keeping_unit' => '300001', 'm_date'=>date("Y-m-d h:i:s"), 'country'=>'SG', 'customer_id'=>$recordId];

                DB::table('mainmerge')->insert($subsampleRecords);
			}
		}
		else
		{
			DB::table($this->table)->insert($this->arr);
		}

		//Looping Data Input Again After Insert

		foreach($this->data_inputan as $ro) {

			$name = $ro['name'];
			if(!$name) continue;

			$inputdata = Request::get($name);

			//Insert Data Checkbox if Type Datatable
			if($ro['type'] == 'checkbox') {
				if($ro['relationship_table']) {
					$datatable = explode(",",$ro['datatable'])[0];
					$foreignKey2 = CRUDBooster::getForeignKey($datatable,$ro['relationship_table']);
					$foreignKey = CRUDBooster::getForeignKey($this->table,$ro['relationship_table']);
					DB::table($ro['relationship_table'])->where($foreignKey,$id)->delete();

					if($inputdata) {
						$relationship_table_pk = CB::pk($ro['relationship_table']);
						foreach($inputdata as $input_id) {
							DB::table($ro['relationship_table'])->insert([
								$relationship_table_pk=>CRUDBooster::newId($ro['relationship_table']),
								$foreignKey=>$id,
								$foreignKey2=>$input_id
								]);
						}
					}

				}
			}

			if($ro['type'] == 'select2') {
				if($ro['relationship_table']) {
					$datatable = explode(",",$ro['datatable'])[0];
					$foreignKey2 = CRUDBooster::getForeignKey($datatable,$ro['relationship_table']);
					$foreignKey = CRUDBooster::getForeignKey($this->table,$ro['relationship_table']);
					DB::table($ro['relationship_table'])->where($foreignKey,$id)->delete();

					if($inputdata) {
						foreach($inputdata as $input_id) {
							$relationship_table_pk = CB::pk($row['relationship_table']);
							DB::table($ro['relationship_table'])->insert([
								$relationship_table_pk=>CRUDBooster::newId($ro['relationship_table']),
								$foreignKey=>$id,
								$foreignKey2=>$input_id
								]);
						}
					}
				}
			}

			if($ro['type']=='child') {

				$tempId = array();
				$name = str_slug($ro['label'],'');
				$columns = $ro['columns'];

				if(Request::get($name.'-'.$columns[0]['name']))
				{
					$count_input_data = count(Request::get($name.'-'.$columns[0]['name']))-1;
					$child_array = [];
					$childtable = CRUDBooster::parseSqlTable($ro['table'])['table'];
					$fk = $ro['foreign_key'];
					$childtablePK = CB::pk($childtable);

					$colMatch = array();
					$row2 = (array) $row;
					foreach ($this->col as $key => $value) {
						$val = $value['name'];
						$colMatch[] = $val;

					}

					$matchRow = [];
					foreach ($colMatch as $key => $field) {
						if(array_key_exists($field, $row2)){
							$matchRow[$field] = $row2[$field];
						}
					}

					for($i=0;$i<=$count_input_data;$i++) {

						$column_data = [];
						// $column_data[$childtablePK] = $lastId;
						$column_data[$fk] = $id;
						foreach($columns as $col) {
							$colname = $col['name'];
							$column_data[$colname] = Request::get($name.'-'.$colname)[$i];
						}

						$child_array[] = $column_data;

						if($child_array[$i]['id'] == NULL){

							if($childtable == 'mainmerge') {
								$customer_array[] = $matchRow;
								$test = (array) $customer_array[$i];
								foreach($child_array as $key => $value)
								{
									$newArray = array_merge($child_array[$key],$test);
								}

								unset($newArray['id']);

								$lastId = CRUDBooster::newId($childtable);
								$newArray['id'] = $lastId;
								date_default_timezone_set("Asia/Kuala_Lumpur");
								$date = date('Y-m-d H:i:s');
								$newArray['m_date'] = $date;

								DB::table($childtable)->insert($newArray);
							}
							else {
								unset($child_array[$i]['id']);
								$lastId = CRUDBooster::newId($childtable);
								$child_array[$i]['id'] = $lastId;
								if($ro['name']=='gigya_children')
								{
									$child_array[$i]['applicationInternalIdentifier'] = $this->generateUid();
									$child_array[$i]['interestCode'] = 'GG_CHILD_MILK_BRAND';

									switch($child_array[$i]['sex'])
									{
										case 'Male'   : $child_array[$i]['sex'] = 1 ; break;
										case 'Female' : $child_array[$i]['sex'] = 2 ; break;
										default       : break;
									}
								}
								if($ro['name']=='gigya_customer_pets')
								{
									$child_array[$i]['applicationInternalIdentifier'] = $this->generateUid();
								}
								if($ro['name']=='careline_detail')
								{
									$child_array[$i]['created_at'] = date("Y/m/d h:i:sa");
									$child_array[$i]['updated_at'] = date("Y/m/d h:i:sa");
								}

								DB::table($childtable)->insert($child_array[$i]);
							}
						}

						$tempId[] = $child_array[$i]['id'];
						unset($child_array[$i]['id']);

						DB::table($childtable)
						->where('id', $tempId[$i])
						->update($child_array[$i]);
					}
				}
			}
		}

		if($this->gigya_based || $this->gigya_customer || $this->thailand_customer)
		{
			$UID      = NULL;
			$regToken = NULL;

			$response = $this->searchViaEmail($this->arr['email']);

			$results = $response['results'];

			if($results[0]["hasFullAccount"])
			{
				$UID = $results[0]['UID'];
			}

	    	if(!isset($UID))
	    	{
	    		$register = $this->initRegistration();

	    		if(is_array($register) && isset($register["regToken"]))
	    		{
		    		$regToken = $register["regToken"];
	    		}
	    	}

	    	$rowArray = $this->arr;
	    	$recordId = $this->arr[$this->primary_key];

	    	$this->synchroToGigya($UID,$regToken,$rowArray['email'],$rowArray,$recordId,$this->arr,$this->table);
		}

		$this->hook_after_add($this->arr[$this->primary_key]);

		$this->return_url = ($this->return_url)?$this->return_url:Request::get('return_url');

		//insert log
		CRUDBooster::insertLog(trans("crudbooster.log_add",['name'=>$this->arr[$this->title_field],'module'=>CRUDBooster::getCurrentModule()->name]));

		if($this->return_url) {
			if(Request::get('submit') == trans('crudbooster.button_save_more')) {
				CRUDBooster::redirect(Request::server('HTTP_REFERER'),trans("crudbooster.alert_add_data_success"),'success');
			}else{
				CRUDBooster::redirect($this->return_url,trans("crudbooster.alert_add_data_success"),'success');
			}

		}else{
			if(Request::get('submit') == trans('crudbooster.button_save_more')) {
				CRUDBooster::redirect(CRUDBooster::mainpath('add'),trans("crudbooster.alert_add_data_success"),'success');
			}else{
				CRUDBooster::redirect(CRUDBooster::mainpath(),trans("crudbooster.alert_add_data_success"),'success');
			}
		}
	}

	public function arrayMappingtoSD($profile, $data){
		$row = new \stdClass();

		foreach ($profile as $key => $value) {
			if($key == 'zip'){
				$row->postcode = $profile['zip'];
			} elseif($key == 'firstName') {
				$row->firstname = $profile['firstName'];
			} elseif($key == 'lastName'){
				$row->lastname = $profile['lastName'];
			} elseif($key == 'phones') {
				$row->mobileno = $profile['phones'][0]['number'];
			} elseif($key == 'address') {
				$row->address1 = $profile['address'];
			} else {
				$row->$key = $profile[$key];
			}
		}

		if(!isset($row->mobileno))
		{
			$row->mobileno = $data['mobile'];
		}

		return $row;
	}

	public function getEdit($id){
		$this->cbLoader();

		if($this->gigya_customer)
		{
			$customer = Customer::find($id);

			try{
				$client = new Client();
				$response = $client->get('https://www.clubillume.com.sg/user/loyalty/'.$customer->email, [
					'headers' => [
				        'x-api-token' => 'WVqcLsu6l9ixSvSAhLPXAxh5nunZa0MVaKU6JP6QVfJDTT7eHMKy595pAMVRCHKQ99dJo6ewca7jncaA',
				    ]
				]);

				$contents = json_decode($response->getBody()->getContents(), true);

				if($contents['status'] == '200')
				{
					$data = $contents['data'];

					if($data['loyalty'])
					{
						$loyaltyPoints = $data['loyalty'];
					}

					if($data['bagtag'])
					{
						$personalisedBagTag = $data['bagtag'];
					}

					$dataMapped = [];
					$subSample  = [];

					foreach($loyaltyPoints as $key => $value)
					{
						switch($key)
						{
							case 'currentPoint': $dataMapped['past_three_months_points'] = $value; break;
							case 'totalPoint'  : $dataMapped['total_lifetime_points']    = $value; break;
							case 'expiryDate'  : $dataMapped['membership_expiry_date']   = $value; break;
							case 'tier'        : $dataMapped['membership_status']        = $value; break;
							default            : break;
						}
					}

					if(isset($dataMapped['past_three_months_points']))
					{
						$dataMapped['points_to_retain_solitaire'] = 1800 - $dataMapped['past_three_months_points'];
					}

					if(isset($loyaltyPoints['user_first_time_solitaire']))
					{
						if($loyaltyPoints['user_first_time_solitaire'] == FALSE)
						{
							$dataMapped['membership_status'] = 'Returning Solitaire';
						}
					}

					DB::table($this->table)
					->where('email', $customer->email)
					->update($dataMapped);

					if(isset($customer->mainMerge))
					{
						$subSample['personalised_tag_name'] = $personalisedBagTag;
						$customer->mainMerge->update($subSample);
					}
				}
			}
			catch (\Exception $e)
            {
                Log::info($e->getMessage());
            }

		}

		$row = DB::table($this->table)->where($this->primary_key,$id)->first();

		// if(isset($row->email) && $this->gigya_based)
		// {
		// 	$response = $this->searchViaEmail($row->email);
		// 	$results  = $response['results'];
		// 	$profile  = $results[0]['profile'];
		// 	$data     = $results[0]['data'];

		// 	if($profile)
		// 	{
		// 		$profile = $this->arrayMappingtoSD($profile, $data);
		// 		foreach ($row as $key1 => $value1) {
		// 			foreach ($profile as $key2 => $value2) {
		// 				if($key2 == $key1){
		// 					$row->$key1 = $profile->$key2;
		// 				}
		// 			}
		// 		}
		// 	}
		// }

		if(!CRUDBooster::isRead() && $this->global_privilege==FALSE || $this->button_edit==FALSE) {
			CRUDBooster::insertLog(trans("crudbooster.log_try_edit",['name'=>$row->{$this->title_field},'module'=>CRUDBooster::getCurrentModule()->name]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans('crudbooster.denied_access'));
		}

		$page_menu       = Route::getCurrentRoute()->getActionName();
		$page_title 	 = trans("crudbooster.edit_data_page_title",['module'=>CRUDBooster::getCurrentModule()->name,'name'=>$row->{$this->title_field}]);
		$command 		 = 'edit';
		Session::put('current_row_id',$id);
		$option_id		 = $this->option_id;
		$option_fields	 = $this->option_fields;
		$table           = $this->table;
		$gigya_customer  = $this->gigya_customer;

		return view('crudbooster::default.form',compact('id','row','page_menu','page_title','command','option_id','option_fields','table', 'gigya_customer'));
	}

	public $countChild = 0;
	public function postEditSave($id) {

		$this->cbLoader();
		$row = DB::table($this->table)->where($this->primary_key,$id)->first();

		if(!CRUDBooster::isUpdate() && $this->global_privilege==FALSE) {
			CRUDBooster::insertLog(trans("crudbooster.log_try_add",['name'=>$row->{$this->title_field},'module'=>CRUDBooster::getCurrentModule()->name]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans('crudbooster.denied_access'));
		}

		$this->validation($id);
		$this->input_assignment($id);

		if (Schema::hasColumn($this->table, 'updated_at'))
		{
		    $this->arr['updated_at'] = date('Y-m-d H:i:s');
		}

		if($this->table == 'cms_users')
		{
			if(isset($this->arr['password']))
			{
				$this->arr['password_updated_at'] = date('Y-m-d H:i:s');
			}

			if(isset($this->arr['status']))
			{
				if($this->arr['status'] == 'Active')
				{
					$this->arr['failed_login_attempts'] = 0;
				}
			}
		}

		$this->hook_before_edit($this->arr,$id);

		//Looping Data Input Again After Insert

		foreach($this->data_inputan as $ro) {

			$name = $ro['name'];

			$type = $ro['type'];

			if(!$name) continue;

			$inputdata = Request::get($name);
			$setInputData[$name] = $inputdata;

			//Insert Data Checkbox if Type Datatable
			if($ro['type'] == 'checkbox') {
				if($ro['relationship_table']) {
					$datatable = explode(",",$ro['datatable'])[0];

					$foreignKey2 = CRUDBooster::getForeignKey($datatable,$ro['relationship_table']);
					$foreignKey = CRUDBooster::getForeignKey($this->table,$ro['relationship_table']);
					DB::table($ro['relationship_table'])->where($foreignKey,$id)->delete();

					if($inputdata) {
						foreach($inputdata as $input_id) {
							$relationship_table_pk = CB::pk($ro['relationship_table']);
							DB::table($ro['relationship_table'])->insert([
								$relationship_table_pk=>CRUDBooster::newId($ro['relationship_table']),
								$foreignKey=>$id,
								$foreignKey2=>$input_id
								]);
						}
					}
				}
			}

			if($ro['type'] == 'select2') {
				if($ro['relationship_table']) {
					$datatable = explode(",",$ro['datatable'])[0];

					$foreignKey2 = CRUDBooster::getForeignKey($datatable,$ro['relationship_table']);
					$foreignKey = CRUDBooster::getForeignKey($this->table,$ro['relationship_table']);
					DB::table($ro['relationship_table'])->where($foreignKey,$id)->delete();

					if($inputdata) {
						foreach($inputdata as $input_id) {
							$relationship_table_pk = CB::pk($ro['relationship_table']);
							DB::table($ro['relationship_table'])->insert([
								$relationship_table_pk=>CRUDBooster::newId($ro['relationship_table']),
								$foreignKey=>$id,
								$foreignKey2=>$input_id
								]);
						}
					}


				}
			}


			if($ro['type']=='child') {
				$tempId = array();
				$name = str_slug($ro['label'],'');
				$columns = $ro['columns'];
				if(Request::get($name.'-'.$columns[0]['name']))
				{
					$count_input_data = count(Request::get($name.'-'.$columns[0]['name']))-1;
					$child_array = [];
					$childtable = CRUDBooster::parseSqlTable($ro['table'])['table'];
					$fk = $ro['foreign_key'];
					$childtablePK = CB::pk($childtable);

					$colMatch = array();
					$row2 = (array) $row;
					foreach ($this->col as $key => $value) {
						$val = $value['name'];
						$colMatch[] = $val;
					}

					$matchRow = [];
					foreach ($colMatch as $key => $field) {
						if(array_key_exists($field, $row2)){
							$matchRow[$field] = $row2[$field];
						}
					}

					for($i=0;$i<=$count_input_data;$i++) {

						$column_data = [];
						$column_data[$fk] = $id;
						foreach($columns as $col) {
							$colname = $col['name'];
							$column_data[$colname] = Request::get($name.'-'.$colname)[$i];
						}

						$child_array[] = $column_data;

						if($ro['name']=='gigya_children')
						{
							switch($child_array[$i]['sex'])
							{
								case 'Male'   : $child_array[$i]['sex'] = 1 ; break;
								case 'Female' : $child_array[$i]['sex'] = 2 ; break;
								default       : break;
							}
						}

						if($child_array[$i]['id'] == NULL){

							if($childtable == 'mainmerge') {
								$customer_array[] = $matchRow;
								$test = (array) $customer_array[$i];
								foreach($child_array as $key => $value)
								{
									$newArray = array_merge($child_array[$key],$test);
								}
								unset($newArray['id']);
								$newArray['mobileno'] = $newArray['phones'];
								$newArray['postcode'] = $newArray['zip'];
								$remove_array = ['phones','zip','careline_max_datecreated','careline_callstatus','careline_currentstatus','careline_telecomaction','mainmerge_max_mdate'];
								$newArray = array_diff_key($newArray, array_flip($remove_array));
								$lastId = CRUDBooster::newId($childtable);
								$newArray['id'] = $lastId;
								date_default_timezone_set("Asia/Kuala_Lumpur");
								$date = date('Y-m-d H:i:s');
								$newArray['m_date'] = $date;

								DB::table($childtable)->insert($newArray);
							}
							else {
								unset($child_array[$i]['id']);
								$lastId = CRUDBooster::newId($childtable);
								if($ro['name']=='gigya_children')
								{
									$child_array[$i]['applicationInternalIdentifier'] = $this->generateUid();
									$child_array[$i]['interestCode'] = 'GG_CHILD_MILK_BRAND';
								}
								if($ro['name']=='gigya_customer_pets')
								{
									$child_array[$i]['applicationInternalIdentifier'] = $this->generateUid();
								}
								if($ro['name']=='careline_detail')
								{
									$child_array[$i]['created_at'] = date("Y/m/d h:i:sa");
									$child_array[$i]['updated_at'] = date("Y/m/d h:i:sa");
								}

								$success = DB::table($childtable)->insert($child_array[$i]);
							}
						}
						$tempId[] = $child_array[$i]['id'];
						unset($child_array[$i]['id']);

						DB::table($childtable)
						->where('id', $tempId[$i])
						->update($child_array[$i]);
					}
				}
			}

		}//end foreach

		if ($setInputData['address2'] == 'NA' || $setInputData['address2'] == 'na') {
			$setInputData['address2'] = '';
		}

		if(isset($this->arr['subsource_id']))
		{
			$this->arr['m_subsource'] = $this->arr['subsource_id'];
			unset($this->arr['subsource_id']);
		}

		DB::table($this->table)->where($this->primary_key,$id)->update($this->arr);

		$UID      = NULL;
		$regToken = NULL;

		if($this->gigya_customer || $this->gigya_based || $this->thailand_customer)
		{
			$response = $this->searchViaEmail($row->email);
			$results = $response['results'];

			if($results[0]["hasFullAccount"])
			{
				$UID = $results[0]['UID'];
			}

			if(!isset($UID))
	    	{
	    		$register = $this->initRegistration();

	    		if(is_array($register) && isset($register["regToken"]))
	    		{
		    		$regToken = $register["regToken"];
	    		}
	    	}

			$this->synchroToGigya($UID,$regToken,$row->email,$setInputData,$id,$this->arr,$this->table);
		}

		$this->hook_after_edit($id);

		$this->return_url = ($this->return_url)?$this->return_url:Request::get('return_url');

		// insert customized cms_log
		if($this->gigya_customer)
		{
			$oldRecord = (array) $row;
			$oldAssignedUser = DB::table('cms_users')->where('id', $oldRecord['userid'])->first();
			if($oldAssignedUser)
			{
				$oldRecord['userid'] = $oldAssignedUser->name;
			}

			$newRecord = $this->arr;
			$email     = $newRecord['email'];

			// Compare the keys and values of two arrays, and return the differences:
			$arrayDiff = array_diff_assoc($newRecord, $oldRecord);

			if(isset($arrayDiff))
			{
				$fields      = array_keys($arrayDiff);
				$description = '<ul>';

				foreach($arrayDiff as $field => $value)
				{
					$label = NULL;

					if($oldRecord[$field] == '')
					{
						$oldRecord[$field] = 'Empty';
					}

					foreach($this->data_inputan as $key => $row)
					{
						if($field == $row['name'])
						{
							$label = $row['label'];

							if ($field == 'userid')
							{
								$assignedUser = DB::table('cms_users')->where('id', $value)->first();
								if($assignedUser)
								{
									$value = $assignedUser->name;
								}
								else
								{
									$value = 'Empty';
								}
							}
						}
						else if($field == 'updated_at')
						{
							$label = 'Updated At';
						}
					}

					$description .= '<li>Update data '.$label.' from '.$oldRecord[$field].' to '.$value.' for '.$email.' at '.CRUDBooster::getCurrentModule()->name.' .'."<br></li>";
				}

				$description .= '</ul>';

				CRUDBooster::insertLog($description);
			}
		}
		else
		{
			//insert log
			CRUDBooster::insertLog(trans("crudbooster.log_update",['name'=>$this->arr[$this->title_field],'module'=>CRUDBooster::getCurrentModule()->name]));
		}


		if($this->return_url) {
			CRUDBooster::redirect($this->return_url,trans("crudbooster.alert_update_data_success"),'success');
		}else{
			if(Request::get('submit') == trans('crudbooster.button_save_more')) {
				CRUDBooster::redirect(CRUDBooster::mainpath('add'),trans("crudbooster.alert_update_data_success"),'success');
			}else{
				CRUDBooster::redirect(CRUDBooster::mainpath(),trans("crudbooster.alert_update_data_success"),'success');
			}
		}
	}

	public function getDelete($id) {
		$this->cbLoader();
		$row = DB::table($this->table)->where($this->primary_key,$id)->first();

		if(!CRUDBooster::isDelete() && $this->global_privilege==FALSE || $this->button_delete==FALSE) {
			CRUDBooster::insertLog(trans("crudbooster.log_try_delete",['name'=>$row->{$this->title_field},'module'=>CRUDBooster::getCurrentModule()->name]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans('crudbooster.denied_access'));
		}

		//insert log
		CRUDBooster::insertLog(trans("crudbooster.log_delete",['name'=>$row->{$this->title_field},'module'=>CRUDBooster::getCurrentModule()->name]));

		$this->hook_before_delete($id);

		if(CRUDBooster::isColumnExists($this->table,'deleted_at')) {
			DB::table($this->table)->where($this->primary_key,$id)->update(['deleted_at'=>date('Y-m-d H:i:s')]);
		}else{
			DB::table($this->table)->where($this->primary_key,$id)->delete();
		}


		$this->hook_after_delete($id);

		$url = g('return_url')?:CRUDBooster::referer();

		CRUDBooster::redirect($url,trans("crudbooster.alert_delete_data_success"),'success');
	}

	public function getDetail($id)	{
		$this->cbLoader();

		$row  = DB::table($this->table)->where($this->primary_key,$id)->first();

		// if(isset($row->email) && $this->gigya_based)
		// {
		// 	$response = $this->searchViaEmail($row->email);

		// 	$results = $response['results'];
		// 	$profile = $results[0]['profile'];
		// 	$data    = $results[0]['data'];
		// 	$UID     = null;

		// 	if(isset($results))
		// 	{
		// 		$profile = $this->arrayMappingtoSD($profile, $data);
		// 		foreach ($row as $key1 => $value1) {
		// 			foreach ($profile as $key2 => $value2) {
		// 				if($key2 == $key1){
		// 					$row->$key1 = $profile->$key2;
		// 				}
		// 			}
		// 		}
		// 	}
		// }

		if(!CRUDBooster::isRead() && $this->global_privilege==FALSE || $this->button_detail==FALSE) {
			CRUDBooster::insertLog(trans("crudbooster.log_try_view",['name'=>$row->{$this->title_field},'module'=>CRUDBooster::getCurrentModule()->name]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans('crudbooster.denied_access'));
		}

		$module     = CRUDBooster::getCurrentModule();

		$page_menu  = Route::getCurrentRoute()->getActionName();
		$page_title = trans("crudbooster.detail_data_page_title",['module'=>$module->name,'name'=>$row->{$this->title_field}]);
		$command    = 'detail';

		Session::put('current_row_id',$id);

		return view('crudbooster::default.form',compact('row','page_menu','page_title','command','id'));
	}

	function csvToArray($filename = '', $delimiter = ',')
	{
	    if (!file_exists($filename) || !is_readable($filename))
	    {
	    	Log::error("Can't find file or not readable");
	    	Log::error($filename);
	    	return false;
	    }

	    $header = null;
	    $data = array();
	    if (($handle = fopen($filename, 'r')) !== false)
	    {
	        while (($row = fgetcsv($handle, 1000, $delimiter)) !== false)
	        {
	        	//Log::error("in CSV");
	        	//Log::error($row);
	            if (!$header)
	            {
	                $header = $row;
	                //Log::error($header);
	            }
	            else
	            {
	            	if (count($row)<count($header))
	            		array_pad($row, count($header), '');
	            	$data[] = array_combine($header, $row);
	            }


	        }
	        fclose($handle);
	    }
	    else
	    {
	    	Log::error("Can't find file");
	    	Log::error($filename);
	    }

	    return $data;
	}

	public function removeMobileNumberDuplication($rows)
	{
		$newRows = [];

		foreach($rows as $key => $value)
		{
			$existingMobileNo = DbtWhatsappNumber::where("mobileno", $value["mobileno"])->first();
			$customerMobileNo = DB::table("customerview")->where("mobileno", $value["mobileno"])->first();
			// $customerMobileNo = Customer::where("mobileno", $value["mobileno"])->first();

			if(isset($existingMobileNo) || isset($customerMobileNo))
			{
				continue;
			}

			$newRows[$key] = $value;
		}

		return $newRows;
	}

	public function getImportData() {
		$this->cbLoader();
		$data['page_menu']       = Route::getCurrentRoute()->getActionName();
		$data['page_title']      = 'Import Data '.$module->name;

		if(Request::get('file') && !Request::get('import')) {
			$file = base64_decode(Request::get('file'));
			$file = trim(str_replace('uploads','app',$file),'/');
			$file = storage_path($file);
			/*$rows = Excel::load($file,function($reader) {
			})->get();*/
			$rows = $this->csvToArray($file);
			//Log::error($rows);
			//Log::error(count($rows));
			//return (count($rows));

			Session::put('total_data_import',count($rows));

			$data_import_column = array();
			foreach($rows as $value) {
				$a = array();
				foreach($value as $k=>$v) {
					$a[] = $k;
				}
				if(count($a)) {
					$data_import_column = $a;
				}
				break;
			}
			//Log::error($data_import_column);

			$table_columns = DB::getSchemaBuilder()->getColumnListing($this->table);
			// Log::error($table_columns);
			// Log::debug(CRUDBooster::myPrivilegeId());

			$data['table_columns'] = $table_columns;
			$data['data_import_column'] = $data_import_column;
		}


		return view('crudbooster::import',$data);
	}

	public function postDoneImport() {
		$this->cbLoader();
		$data['page_menu']       = Route::getCurrentRoute()->getActionName();
		$data['page_title']      = trans('crudbooster.import_page_title',['module'=>$module->name]);
		Session::put('select_column',Request::get('select_column'));

		return view('crudbooster::import',$data);
	}

	public function postDoImportChunk() {
		$this->cbLoader();
		if(CRUDBooster::myPrivilegeId() == 1){
			$file_md5 = md5(Request::get('file'));
			if(Request::get('file') && Request::get('resume')==1) {
				$total = Session::get('total_data_import');
				$prog = intval(Cache::get('success_'.$file_md5)) / $total * 100;
				$prog = round($prog,2);
				if($prog >= 100) {
					Cache::forget('success_'.$file_md5);
				}
				return response()->json(['progress'=> $prog, 'last_error'=>Cache::get('error_'.$file_md5) ]);
			}
			$select_column = Session::get('select_column');
			$select_column = array_filter($select_column);
			$table_columns = DB::getSchemaBuilder()->getColumnListing($this->table);
			$file = base64_decode(Request::get('file'));
			$file = trim(str_replace('uploads','app',$file),'/');
			$file = storage_path($file);
			$rows = Excel::load($file,function($reader) {
			})->get();
			$has_created_at = false;
			if(Schema::hasColumn($this->table,'created_at')) {
				$has_created_at = true;
			}
			$data_import_column = array();
			foreach($rows as $value) {
				$a = array();
				foreach($select_column as $sk => $s) {
					$colname = $table_columns[$sk];
					if(CRUDBooster::isForeignKey($colname)) {
						//Skip if value is empty
						if($value->$s == '') continue;
						if(intval($value->$s)) {
							$a[$colname] = $value->$s;
						}else{
							$relation_table = CRUDBooster::getTableForeignKey($colname);
							$relation_moduls = DB::table('cms_moduls')->where('table_name',$relation_table)->first();
							$relation_class = __NAMESPACE__ . '\\' . $relation_moduls->controller;
							if(!class_exists($relation_class)) {
								$relation_class = '\App\Http\Controllers\\'.$relation_moduls->controller;
							}
							$relation_class = new $relation_class;
							$relation_class->cbLoader();
							$title_field = $relation_class->title_field;
							$relation_insert_data = array();
							$relation_insert_data[$title_field] = $value->$s;
							if(CRUDBooster::isColumnExists($relation_table,'created_at')) {
								$relation_insert_data['created_at'] = date('Y-m-d H:i:s');
							}
							try{
								$relation_exists = DB::table($relation_table)->where($title_field,$value->$s)->first();
								if($relation_exists) {
									$relation_primary_key = $relation_class->primary_key;
									$relation_id = $relation_exists->$relation_primary_key;
								}else{
									$relation_id = DB::table($relation_table)->insertGetId($relation_insert_data);
								}
								$a[$colname] = $relation_id;
							}catch(\Exception $e) {
								exit($e);
							}
						} //END IS INT
					}else{
						$a[$colname] = $value->$s;
					}
				}

				$has_title_field = true;

				foreach($a as $k=>$v) {
					if($k == $this->title_field && $v == '') {
						$has_title_field = false;
						break;
					}
				}

				if($has_title_field==false) continue;

				if($has_created_at) {
					$a['created_at'] = date('Y-m-d H:i:s');
				}

				try{
					if($this->import_mobile_number)
					{
						if(isset($a['mobileno']) && isset($a['email']))
						{
							$email    = $a['email'];
							$record   = DB::table($this->table)->where("email", $email);

							if($record)
							{
								$update = $record->update(['mobileno' => $a['mobileno']]);
							}
						}
					}
					else
					{
						DB::table($this->table)->insert($a);
					}
					Cache::increment('success_'.$file_md5);
				}catch(\Exception $e) {
					$e = (string) $e;
					Cache::put('error_'.$file_md5,$e,500);
				}
			}
		} else {
			$file_md5 = md5(Request::get('file'));
			Log::debug($file_md5);
			Cache::add('success_'.$file_md5, 0, 60);

			if(Request::get('file') && Request::get('resume')==1) {
				$total = Session::get('total_data_import');
				$prog = intval(Cache::get('success_'.$file_md5)) / $total * 100;
				$prog = round($prog,2);
				if($prog >= 100) {
					Cache::forget('success_'.$file_md5);
				}
				return response()->json(['progress'=> $prog, 'last_error'=>Cache::get('error_'.$file_md5) ]);
			}

			$select_column = Session::get('select_column');
			$select_column = array_filter($select_column);
			$table_columns = DB::getSchemaBuilder()->getColumnListing($this->table);

			$file = base64_decode(Request::get('file'));
			$file = trim(str_replace('uploads','app',$file),'/');
			$file = storage_path($file);

			$rows = $this->csvToArray($file);

			if(isset($rows) && $this->table == 'dbt_whatsapp_numbers')
			{
				DB::statement("DROP VIEW IF EXISTS CustomerView");

				DB::statement("CREATE VIEW CustomerView AS
						SELECT c.firstname, c.lastname, c.email, c.mobileno, g.preference_name, g.customer_id
						FROM customer c
						INNER JOIN gigya_preferences g
						ON g.customer_id = c.id"
					 );

				$rows = $this->removeMobileNumberDuplication($rows);

				DB::statement("DROP VIEW IF EXISTS CustomerView");
			}
			
			$f = $this->import_consignment;

			$has_created_at = false;
			if(CRUDBooster::isColumnExists($this->table,'created_at')) {
				$has_created_at = true;
			}

			$data_import_column = array();
			$uploadNotUpdated = [];

			// return response()->json(['select_column'=>$select_column, 'table_columns' => $table_columns]);

			if($this->sfmc_alert)
			{
				$batch = DB::table($this->table)->max('batch');
				$batch++;
			}

			foreach($rows as $value) {

				$a = array();
				foreach($select_column as $sk => $s) {
					$colname = $table_columns[$sk];
					$a[$colname] = $value[$s];
				}

				$has_title_field = true;
				foreach($a as $k=>$v) {
					if($k == $this->title_field && $v == '') {
						$has_title_field = false;
						break;
					}
				}

				if($has_title_field==false) continue;

				try{
					$f = $this->import_consignment;
					if ($f == FALSE)
					{
						if($has_created_at) {
							$a['created_at'] = date('Y-m-d H:i:s');
						}

						if($this->import_offline)
						{
							if(!isset($a['m_date']))
							{
								$a['m_date'] = date("Y-m-d H:i:s");
							}
						
							if(isset($a['childdob']))
							{
								$dateString = str_replace('/', '-', $a['childdob']); 
								$a['childdob'] = date("Y-m-d", strtotime($dateString));
							}

							$existingRecords = DB::table($this->table)->where("email", $a["email"])->where("m_product", $a["m_product"])->get()->toArray();

							if(count($existingRecords) >= 2)
							{
								$latestRecord = end($existingRecords);

								DB::table($this->table)
                                ->where("id", $latestRecord->id)
                                ->update($a);

                                continue;
							}
						}

						if($this->sfmc_alert)
						{
							$a['batch'] = $batch;
						}

						DB::table($this->table)->insert($a);

						$uploadStatus = 'Successful';
					}
					else
					{
						if ((isset($a['m_product'])) || (isset($a['m_date'])) || (isset($a['email'])) || (isset($a['mobileno'])) || (isset($a['childname'])) || (isset($a['childdob'])) )
						{
							if (($a['consigmentno'] != '')||($a['returnreason'] != '')||($a['batchno'] != '')||$a['delivery_status'] != '')
							{
								$a['order_number']    = $value['order_number'];
								$a['delivery_status'] = $value['delivery_status'];

								$recordID = (int) substr($a['order_number'],3);
								$record   = DB::table($this->table)->where("id", $recordID);

								if($record)
								{
									$countCheckDB = $record->update(['consigmentno' => $a['consigmentno'],'batchno' => $a['batchno'],'returnreason'=> $a['returnreason'], 'delivery_status' => $a['delivery_status']]);
								}

								if($countCheckDB == 0){
									$uploadNotUpdated[] = $a;
								}

								$uploadStatus = 'Successful';
							}
						} else {
							$uploadStatus = "Fail to match record";
						}
					}
					Cache::increment('success_'.$file_md5);
				}catch(\Exception $e) {
					$e = (string) $e;
					$uploadStatus = 'Failed';
					Log::error('Error'.$e);
					Cache::put('error_'.$file_md5,$e,500);
				}
			}

			Log::debug($uploadNotUpdated);

			DB::table('upload_logs')->insert([
				[
					'userid' => CRUDBooster::myId(),
					'status' => $uploadStatus,
					'created' => Carbon\Carbon::now()
				]
			]);
		}
		return response()->json(['status'=>true]);
	}

	public function postDoUploadImportData() {
		$this->cbLoader();
		if (Request::hasFile('userfile'))
		{
			$file = Request::file('userfile');
			$ext  = $file->getClientOriginalExtension();


			$validator = Validator::make([
				'extension'=>$ext,
				],[
				//'extension'=>'in:xls,xlsx,csv'
				'extension'=>'in:csv'
				]);

		    if ($validator->fails())
		    {
		        $message = $validator->errors()->all();
		        return redirect()->back()->with(['message'=>implode('<br/>',$message),'message_type'=>'warning']);
		    }

			//Create Directory Monthly
			Storage::makeDirectory(date('Y-m'));


			//Move file to storage
			$filename = md5(str_random(5)).'.'.$ext;
			$url_filename = '';
			if($file->move(storage_path('app'.DIRECTORY_SEPARATOR.date('Y-m')),$filename)) {
				$url_filename = 'uploads/'.date('Y-m').'/'.$filename;
			}

			$url = CRUDBooster::mainpath('import-data').'?file='.base64_encode($url_filename);

			return redirect($url);
		}else{
			return redirect()->back();
		}
	}

	public function postActionSelected() {
		$this->cbLoader();
		$id_selected = Request::input('checkbox');
		$button_name = Request::input('button_name');

		if($button_name == 'delete') {
			if(!CRUDBooster::isDelete()) {
				CRUDBooster::insertLog(trans("crudbooster.log_try_delete_selected",['module'=>CRUDBooster::getCurrentModule()->name]));
				CRUDBooster::redirect(CRUDBooster::adminPath(),trans('crudbooster.denied_access'));
			}

			$this->hook_before_delete($id_selected);
			$tablePK = CB::pk($this->table);
			if(CRUDBooster::isColumnExists($this->table,'deleted_at')) {

				DB::table($this->table)->whereIn($tablePK,$id_selected)->update(['deleted_at'=>date('Y-m-d H:i:s')]);
			}else{
				DB::table($this->table)->whereIn($tablePK,$id_selected)->delete();
			}
			CRUDBooster::insertLog(trans("crudbooster.log_delete",['name'=>implode(',',$id_selected),'module'=>CRUDBooster::getCurrentModule()->name]));

			$this->hook_after_delete($id_selected);

			$message = trans("crudbooster.alert_delete_selected_success");
			return redirect()->back()->with(['message_type'=>'success','message'=>$message]);
		}

		$action = str_replace(['-','_'],' ',$button_name);
		$action = ucwords($action);
		$type = 'success';
		$message = trans("crudbooster.alert_action",['action'=>$action]);
		$custom_message="";

		if($this->actionButtonSelected($id_selected,$button_name,$custom_message) === false) {
			echo "Custom Message:".$custom_message."<br>";
		    $message = !empty($this->alert['message']) ? (($custom_message=="")?$this->alert['message']:$custom_message) : 'Error';
		    //$message = !empty($this->alert['message']) ? $this->alert['message'] : 'Error';

		    $type = !empty($this->alert['type']) ? $this->alert['type'] : 'danger';
		}
		else
			$message = ($custom_message=="")?$message:$custom_message;

		return redirect()->back()->with(['message_type'=>$type,'message'=>$message]);
	}

	public function getDeleteImage() {
		$this->cbLoader();
		$id     = Request::get('id');
		$column = Request::get('column');

		$row    = DB::table($this->table)->where($this->primary_key,$id)->first();

		if(!CRUDBooster::isDelete() && $this->global_privilege==FALSE) {
			CRUDBooster::insertLog(trans("crudbooster.log_try_delete_image",['name'=>$row->{$this->title_field},'module'=>CRUDBooster::getCurrentModule()->name]));
			CRUDBooster::redirect(CRUDBooster::adminPath(),trans('crudbooster.denied_access'));
		}

		$row = DB::table($this->table)->where($this->primary_key,$id)->first();

		$file = str_replace('uploads/','',$row->{$column});
		if(Storage::exists($file)) {
        	Storage::delete($file);
       	}

       	DB::table($this->table)->where($this->primary_key,$id)->update([$column=>NULL]);

		CRUDBooster::insertLog(trans("crudbooster.log_delete_image",['name'=>$row->{$this->title_field},'module'=>CRUDBooster::getCurrentModule()->name]));

		CRUDBooster::redirect(Request::server('HTTP_REFERER'),trans('crudbooster.alert_delete_data_success'),'success');
	}

	public function postUploadSummernote() {
		$this->cbLoader();
		$name = 'userfile';
		if (Request::hasFile($name))
		{
			$file = Request::file($name);
			$ext  = $file->getClientOriginalExtension();

			//Create Directory Monthly
			Storage::makeDirectory(date('Y-m'));

			//Move file to storage
			$filename = md5(str_random(5)).'.'.$ext;
			if($file->move(storage_path('app'.DIRECTORY_SEPARATOR.date('Y-m')),$filename)) {
				echo asset('uploads/'.date('Y-m').'/'.$filename);
			}
		}
	}

	public function postUploadFile() {
		$this->cbLoader();
		$name = 'userfile';
		if (Request::hasFile($name))
		{
			$file = Request::file($name);
			$ext  = $file->getClientOriginalExtension();

			//Create Directory Monthly
			Storage::makeDirectory(date('Y-m'));

			//Move file to storage
			$filename = md5(str_random(5)).'.'.$ext;
			if($file->move(storage_path('app'.DIRECTORY_SEPARATOR.date('Y-m')),$filename)) {
				echo 'uploads/'.date('Y-m').'/'.$filename;
			}
		}
	}
	//&$custom_message=""
	public function actionButtonSelected($id_selected,$button_name,&$custom_message = "") {
    }

	public function hook_query_index(&$query) {
	}

	public function hook_row_index($index,&$value) {
    }

	public function hook_before_add(&$arr) {
	}

	public function hook_after_add($id) {
	}

	public function hook_before_edit(&$arr,$id) {
	}

	public function hook_after_edit($id) {
	}

	public function hook_before_delete($id) {
	}

	public function hook_after_delete($id) {
	}

}
