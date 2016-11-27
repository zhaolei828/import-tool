<#assign base=request.getContextPath()>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="author" content="fyunli">

    <base id="base" href="${base}">
    <title>工具 - 导入excel</title>

    <!-- Bootstrap core CSS -->
    <link href="//cdn.jsdelivr.net/bootstrap/3.3.6/css/bootstrap.min.css" rel="stylesheet">
    <link href="${base}/css/main.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
    <script src="//cdn.jsdelivr.net/html5shiv/3.7.3/html5shiv.min.js"></script>
    <script src="//cdn.jsdelivr.net/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
</head>

<body>

<!-- Begin page content -->
<div class="container">
    <div class="page-header">
        <h1>导入excel</h1>
    </div>

    <form class="form-horizontal" action="do_import" method="post">
        <div class="form-group">
            <label for="inputEmail3" class="col-sm-2 control-label">Doc文件目录</label>
            <div class="col-sm-10">
                <input type="text" class="form-control" placeholder="Text input" name="docPath">
            </div>
        </div>
        <div class="form-group">
            <div class="col-sm-offset-2 col-sm-10">
                <button type="submit" class="btn btn-default">提交</button>
            </div>
        </div>
    </form>
    <h3>导入须知</h3>
    <ul>
        <li>
            <p class="lead">注意：</p>
            <ol>
                <li><p>所有的序号不要用word自带的序号，自己手打</p></li>
                <li><p>换行要用回车，不要用下箭头</p></li>
            </ol>
        </li>
    </ul>
    <ul>
        <li>
            <p class="lead">题干查找规则</p>
            <ul>
                <li>
                    <p>数字+.|． <strong>并且</strong> 前面紧挨的元素是空行。</p>
                <li><p>题号不要使用word的序列格式，手动输入即可</p></li>
                </li>
            </ul>
        </li>
        <li>
            <p class="lead">题干组合规则</p>
            <ol>
                <li><p>题干</p></li>
                <li><p>题干与选项之间的内容</p></li>
                <li><p>题干与小题之间的内容</p></li>
            </ol>
        </li>
        <li>
            <p class="lead">找答案文档规则</p>
            <ol>
                <li><p>必须是docx</p></li>
                <li>
                    <p>题目和答案分开在两个word文档中，其中一个文档名称为“**试题”，则这个文档为试题文档，里面都是试题；另一个文档名称为“**答案”或“**答案**”。</p>
                    <p>所以用 “试题”去截取题目的那个文档名，然后去找对应的答案文档。</p>
                </li>
            </ol>
        </li>
        <li>
            <p class="lead">大题查找规则</p>
            <ul><li><p>(一|二|三|四|五|六|七|八|九|十)、</p></li></ul>
        </li>
    </ul>
</div>

<footer class="footer">
    <div class="container">
        <p class="text-muted">&copy;2016 zhaolei</p>
    </div>
</footer>


<!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
<script src="//cdn.jsdelivr.net/ie10-viewport/1.0.0/ie10-viewport.min.js"></script>
<script src="//cdn.jsdelivr.net/jquery/1.12.1/jquery.min.js"></script>
<script src="//cdn.jsdelivr.net/bootstrap/3.3.6/js/bootstrap.min.js"></script>
</body>
</html>