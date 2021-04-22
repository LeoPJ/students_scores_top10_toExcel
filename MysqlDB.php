<?php
header("Content-type:text/html;charset=utf8");
class MysqlDB{
    private $host;//localhost 127.0.0.1 有一个断网后不能使用
    private $username;//用户名
    private $password;//密码
    private $dbname;//数据库名
    private $charset;//编码格式
    private $port; //端口号
    private static $instance;//存放实例
    private $link;//存放链接
    //获取实例静态方法
    public static function getInstance($config){
        //判断是否已经实例化MysqlDB类
        if(!self::$instance){
            //如果没有实例化就实例化
            self::$instance = new self($config);
        }
        //如果已经实例化就返回
        return self::$instance;
    }
    //构造方法
    private function __construct($config) {
        
        $this->host = isset($config['host'])?$config['host']:'*';//判断是否写了服务器地址
        $this->username = isset($config['username'])?$config['username']:'*';//判断是否写了用户名
        $this->password = isset($config['password'])?$config['password']:'*';//判断是否写了密码
        $this->dbname = isset($config['dbname'])?$config['dbname']:'visiting';//判断是否写了数据库名
        $this->charset = isset($config['charset'])?$config['charset']:'charset';//判断是否设置编码格式
        $this->port = isset($config['port'])?$config['port']:'*';//判断是否写了端口号
        
        self::$instance = $this->connect();//获取连接
        $this->setCharset($this->charset);//设置编码格式
        $this->selectDB($this->dbname);//选择数据库
    } 
    //获取连接的方法
    private  function connect(){
        $this->link = mysqli_connect($this->host, $this->username, $this->password, $this->dbname, $this->port);
    }
    //设置编码的方法
    private function setCharset($charset){
        mysqli_set_charset($this->link, $charset);
    }
    private function selectDB($dbname){
        mysqli_select_db($this->link, $dbname);
    }
    //执行查询的方法
    function query($sql){
        //判断是否查询结果是否有误
        if(!$result=mysqli_query($this->link, $sql)){
            echo "执行错误<br>";
            echo "错误的位置是".mysqli_error()."<br>";
            echo "错误的sql是".$sql;
        }
        //没问题就返回结果
        return $result;
    }
    //获取多行多列数据的方法
    function getAll($sql, $fileheader){
        $result = $this->query($sql);
        $arr = array();
        $arr[] = $fileheader;
        if($result !=FALSE){
            while ($list = mysqli_fetch_assoc($result)){
                $arr[] = $list;
            }
            return $arr;
        }
    }
    //获取单行多列的数据
    function getRow($sql){
        $result = $this->query($sql);
        if($result !=FALSE){
            return mysqli_fetch_assoc($result);
        }else{
            return FALSE;
        }
    }
    //获取单行单列的数据
    function getOne($sql){
        $result = $this->query($sql);
        $list = mysqli_fetch_row($result);
        if($result === FALSE){
            return FALSE;
        }
        return $list[0];
    }
}

