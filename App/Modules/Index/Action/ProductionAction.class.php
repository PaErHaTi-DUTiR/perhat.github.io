<?php
/* 前台产品控制器 */
class ProductionAction extends CommonAction{
	/* 模版视图 */
    public function product_info(){
    	$this->display();
    }

	public function index(){
        $db	= M('production');
	    //360道地毯数据提取
	    import('ORG.Util.Page');// 导入分页类
	    $where_360['type'] = '360道地毯';    
		$count_360      = $db->where($where_360)->count();// 查询满足要求的总记录数
		$page_360       = new Page($count_360,6);// 实例化分页类 传入总记录数和每页显示的记录数
		$show_360       = $page_360->show();// 分页显示输出
		// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
		$list_360 = $db->order('addtime DESC')->limit($page_360->firstRow.','.$page_360->listRows)->where($where_360)->select();
		$this->assign('list_360',$list_360);// 赋值数据集
		$this->assign('page_360',$show_360);// 赋值分页输出      //输出页码

        //540道地毯数据提取
	    $where_540['type'] = '540道地毯';    
		$count_540      = $db->where($where_540)->count();// 查询满足要求的总记录数
		$page_540       = new Page($count_540,6);// 实例化分页类 传入总记录数和每页显示的记录数
		$show_540       = $page_540->show();// 分页显示输出
		// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
		$list_540 = $db->order('addtime DESC')->limit($page_540->firstRow.','.$page_540->listRows)->where($where_540)->select();
		$this->assign('list_540',$list_540);// 赋值数据集
		$this->assign('page_540',$show_540);// 赋值分页输出      //输出页码

		//720道地毯数据提取
	    $where_720['type'] = '720道地毯';    
		$count_720      = $db->where($where_720)->count();// 查询满足要求的总记录数
		$page_720       = new Page($count_720,6);// 实例化分页类 传入总记录数和每页显示的记录数
		$show_720       = $page_720->show();// 分页显示输出
		// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
		$list_720 = $db->order('addtime DESC')->limit($page_720->firstRow.','.$page_720->listRows)->where($where_720)->select();
		$this->assign('list_720',$list_720);// 赋值数据集
		$this->assign('page_720',$show_720);// 赋值分页输出      //输出页码

        //sichou道地毯数据提取
	    $where_sichou['type'] = '丝绸地毯';    
		$count_sichou      = $db->where($where_sichou)->count();// 查询满足要求的总记录数
		$page_sichou       = new Page($count_sichou,6);// 实例化分页类 传入总记录数和每页显示的记录数
		$show_sichou       = $page_sichou->show();// 分页显示输出
		// 进行分页数据查询 注意limit方法的参数要使用Page类的属性
		$list_sichou = $db->order('addtime DESC')->limit($page_sichou->firstRow.','.$page_sichou->listRows)->where($where_sichou)->select();
		$this->assign('list_sichou',$list_sichou);// 赋值数据集
		$this->assign('page_sichou',$show_sichou);// 赋值分页输出      //输出页码

		$this->display();
	}
}
?>