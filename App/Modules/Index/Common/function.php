<?php
/*  打印函数*/
function p($arr){
echo '<pre>';
return print_r($arr);
echo '</pre>';


}
/*
 *递归重组节点信息为多维数组
 *$node要处理的多维数组
 *$pid父级ID	
 */
function node_merge($node,$access=null,$pid=0){
	//print_r($node);die;
	$arr=array();
	foreach($node as $v){
	//print_r($v);die;
		if(is_array($access)){
			$v['access']=in_array($v['id'],$access)?1:0;
		
		}
		if($v['pid']==$pid){
			$v['child']=node_merge($node,$access,$v['id']);
			$arr[]=$v;
		
		}
	}
	return $arr;
}//end


?>