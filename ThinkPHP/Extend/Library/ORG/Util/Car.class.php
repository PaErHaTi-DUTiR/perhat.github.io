<?php
class  Car{
	//定义商品数组存放数组
	protected $item=array();
	public function Addgoods($goods_id,$name,$num,$price){
		$this->item['id']=$goods_id;
		$this->item['num']=$num;
		$this->item['name']=$name;
		$this->item['price']=$price;
		if(!$_SESSION["myCart"])$_SESSION["myCart"]=array();
		if(array_key_exists($goods_id,$_SESSION["myCart"])){
	  		$_SESSION["myCart"][$goods_id]['num']+=$num;
		}else{
			$_SESSION["myCart"][$goods_id]=$this->item;
		}
	}
	public function delGoods($goods_id){
		if(in_array($goods_id,$_SESSION["myCart"][$goods_id])){
			unset($_SESSION["myCart"][$goods_id]);
			}else{
			return false;	
			}
	}
	//计算总价格
	public function Getprice(){
		$price=0;
		$dd=$_SESSION["myCart"];
		foreach ($dd as $vv){
			$price+=$vv['price']*$vv['num'];
			}
		return $price;
	}
	//计算总商品数量
	
	public function getNums(){
		$pnums=0;
		$dd=$_SESSION["myCart"];
		foreach ($dd as $vv){
			$pnums+=$vv['num'];
			}
		return $pnums;
	}
	
	
	//减少一件商品
	
	public function reduceGoods($goods_id){
		$current=$_SESSION["myCart"][$goods_id]['num']-=1;;
		return $current; 
		}
	
	//减加一件商品
	
	public function jiaGoods($goods_id){
		$current=$_SESSION["myCart"][$goods_id]['num']+=1;;
		return $current; 
		}
	
	//一种商品的总金额
	public function getHeji($goods_id){
		
		$priceTotal=$_SESSION["myCart"][$goods_id]['num']*$_SESSION["myCart"][$goods_id]['price'];
		return $priceTotal; 
		} 
	
	
}