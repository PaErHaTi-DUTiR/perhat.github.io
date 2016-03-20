<?php
/* 景点控制器 */
class MessageAction extends CommonAction{
	/* 列表视图 */
	public function index(){
	   
	   $this->display();
	}
    
    public function booking(){
    	if($this->isPost()){
			if(!empty($_POST)){
				$data=array(
					//'type'=>$this->_post('type'),
					'note'=>'留言：'.$this->_post('note'),
					'type'=>$this->_post('type'),
					'contact'=>$this->_post('name'),
					'tel'=>$this->_post('phone'),
					'email'=>$this->_post('email'),
					'p_num'=>$this->_post('num'),
					'addtime'=>time(),
					'booknum'=>date('YmdHis',time()).rand('0','99'),
					'bookstatus'=>0,
				
				);
				if(M('book')->add($data)){
					$this->success('预定成功');
				}else{
					$this->error('预定失败');
				}
			}
		
		}
    }	
}
?>