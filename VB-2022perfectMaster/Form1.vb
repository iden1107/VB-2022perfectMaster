Public Class Form1

End Class

Namespace Section3

    '●3.5.1 表示の変換
    Public Sub trem5()
        'データ型変換
        Dim x As String = "true"
        x = CBool(x)

        '通貨変換
        FormatCurrency(10000) '\10,000と変換される

        '0埋め
        Format(55, "0000") '0055

        '小数点表記
        Format(123.34, "0.00") '123.450
        Format(123.34, "0.0") '123.5 この場合四捨五入が発生するので注意

        Format(59800, "\\#,#") '\59.800
        Format(59800, "$#,#") '$59.800
    End Sub

    '●3.5.4 イベントハンドラ
    'フォーカスによりイベントの種類はEnter,leave,validating,validatedなどある
    'フォーカスが当たる時のEnterはあまり使わない
    '外れる際は発生タイミングがleavev→validating→validatedという順番がある

    'leave
    Private Sub subA(sender As Object, e As CancelEventArgs) Handles TextBox1.leave
        Debug.WriteLine("1番最初に発生")
    End Sub

    'validating
    Private Sub subA(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        Debug.WriteLine("2番に発生")
    End Sub

    'validated
    Private Sub subA(sender As Object, e As CancelEventArgs) Handles TextBox1.validated
        Debug.WriteLine("3番に発生")
    End Sub

    'このイベントは入力チェックの際に有効的に使用できる
    '例えばフォーカスが外れた時にテキストボックスの値をチェックするなど
    '以下のコードはtextboxに"hoge"と入力しない限りフォーカスが留まる処理
    '今回は後続にvalidatedイベントでmagboxを開く処理を用意している
    'このコードはe.Cancelがあるため"hoge"と入力しない限り（validatedイベントも含めて）次の処理に移行しない
    '言い換えれば入力が正確でない限りフォーカスが戻り続けるので入力チェックとしては思惑通り機能しているといえる
    Private Sub subA(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        If Not TextBox1.text = "hoge" Then
            MsgBox("hogeと入力してください")
            e.Cancel = True
        End If
    End Sub

    Private Sub subA(sender As Object, e As CancelEventArgs) Handles TextBox1.Validated 'hogeと入力しない限り発生しない
        MsgBox("次の処理")
    End Sub

    'しかし後続してほしい処理がある場合、それでは次の処理に移れない
    '（例えば入力が"hoge"以外だったとしも後続してほしい処理がある場合）
    'その場合次の処理(ここではButton1押下とする)の(button)プロパティのCausesValidation をFalse にすると
    'その処理を実行するときだけはValidatingイベントとValidatedイベントが発生しなくなる
    'なので「このボタンを押すときだけは入力チェックを機能させたくない」などという場合に有効

    'それでも例外的にMe.closeと[×]で閉じる処理だけはValidatingが発生してしまい入力チェックが機能してしまう
    'それでは閉じる時には入力値がおかしいと入力チェックずっと引っかかって閉じれないので挙動とてしは不都合
    'これの解決策としてleaveイベントとfocus()で代替できそうだが、focus()のせいで他の処理ができなくなってしまう
    'validatingとe,cancelを使って入力チェックを機能させつつ、閉じるときは入力チェックを機能させない
    'ようにするために①空文字チェック②空文字リセット　を設けることで両立できる（この場合入力値が毎回クリアされるのが難点）
    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        If TextBox1.Text = "" Then Return
        If TextBox1.Text <> "hoge" Then
            MsgBox("hogeと入力してください")
            TextBox1.Text = ""
            e.Cancel = True
        End If
    End Sub

    'やや冗長的ではあるが別法として遷移先のActiveControl.CausesValidationの真偽値を先に判断する方法がある
    'この場合は入力値を残せる
    'だが[×]では入力チェックが機能されたまま…
    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        If ActiveControl.CausesValidation = False Then
            Exit Sub
        End If
        If TextBox1.Text <> "hoge" Then
            MsgBox("hogeと入力してください")
            e.Cancel = True
        End If
    End Sub


    '●3.7.3 値渡し（byVal）と参照渡し（ByRef）
    Public Sub trem7_3()
        '値渡しは元の値がコピーされる。基本はこちらを使用、省略可能。
        '参照渡しは元の値が変わってしまう。
    End Sub

    '●7.4 デフォルト値
    Public Sub hoge(Optional arg As Integer = 100)
        '引数がなければ100を2倍にする
        Debug.WriteLine(arg * 2)
    End Sub

    '●3.7.5 可変長引数
    Public Sub hoge(ParamArray args() As Integer)
        For Each arg In args
            Debug.WriteLine(arg)
        Next
    End Sub

    hoge(1,2,3,4,5) 'いくらでも引数を取れる
End Namespace


Namespace Section4


    '●4.1.0 アクセス修飾子
    'Private クラス内のみok
    'Protected 継承したクラスからならok
    'Friend 同一プロジェクトからならok
    'Public どこからでもok
    '※修飾子なしならPrivate

    '●フィールドの宣言はアンダーバーを推奨
    Public Class Hoge
        'Dimを使わずアクセス修飾子(Private)を使いカプセル化する
        Private _name As String
        Private _age As Integer

    End Class

    '●プロパティ
    'privateのフィールドは アクセスできないのでプロパティを通してアクセスする
    'getは読み取り専用、setは書き込み専用
    'getプロシージャとsetプロシージャはPulbicを使う（でないとアクセスできない）
    Public Class Hoge
        Private _name As String
        Private _age As Integer

        '_nameの参照と設定するプロパティ
        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        '_ageの参照と設定をするプロパティ
        Public Property Age() As Integer
            Get
                Return _age
            End Get
            Set(value As Integer)
                _age = value
            End Set
        End Property
    End Class


    '●メソッド
    Public Class Hoge
        Private _name As String
        Private _age As Integer

        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        Public Property Age() As Integer
            Get
                Return _age
            End Get
            Set(value As Integer)
                _age = value
            End Set
        End Property

        'メソッドを作成
        Public Function foo() As Integer
            Debug.WriteLine("名前は" + Name + "年齢は" + Age.ToString())
        End Function
    End Class

    'インスタンス化
    Dim person As New Hoge
    person.Name = "ヤマダ"
    person.Age = 28
    person.foo()

    'インスタンス化と同時に初期化する方法
    Dim preson2 As New Hoge With {
        .Name = "田中", '！カンマが必要
        .Age = 46
    }
    person2.foo()

    '●プロパティのチェック機構
    '例えば名前は空白以外、　年齢は0以上でなければならないという条件をつけることができる
    Public Class Hoge
        Private _name As String
        Private _age As Integer

        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal value As String)
                If value = "" Then
                    MsgBox("名前を入力してください")
                Else
                    _name = value
                End If
            End Set
        End Property

        Public Property Age() As Integer
            Get
                Return _age
            End Get
            Set(value As Integer)
                If value < 0 Then
                    MsgBox("0以上で入力してください")
                Else
                    _age = value
                End If
            End Set
        End Property

        Public Sub foo()
            Debug.WriteLine("名前は" + Name + "年齢は" + Age.ToString())
        End Sub
    End Class

    '●Getしかない時はReadOnlyを付ける（読み取り専用）
    Public Class Hoge
        Private _name As String

        Public ReadOnly Property Name() As String
            Get
                Return _name
            End Get
        End Property
    End Class


    '●Setしかない時はWriteOnlyを付ける（書き込み専用）
    Public Class Hoge
        Private _name As String

        Public WriteOnly Property Name() As String
            Set(value As String)
                _dname = value
            End Set
        End Property
    End Class

    '●フィールドが配列（インデックスつきプロパティ）
    Public Class Hoge
        '配列のフィールド
        Private _arrs() As Integer

        Public Property Arrs(index As Integer) As Integer
            Get
                Return _arrs(index)
            End Get
            Set(value As Integer)
                _arrs(index) = value
            End Set
        End Property

        '配列をループで出力するメソッド
        Public Sub foo()
            For Each Arr In _arrs
                Debug.WriteLine(Arr)
            Next
        End Sub
    End Class

    'インスタンス化
    Dim Money As New Hoge
    Money.Arrs(0) = 1000
    Money.Arrs(1) = 5000
    Money.Arrs(2) = 10000
    Money.foo()


    '●4.1.8 自動実装プロパティ

    '自動的にprivateなフィールドを生成し、getとsetが不要になりシンプルになる
    'これだけでgetとsetが内部的に自動生成されている
    Public Class Hoge
        '※Publicになっている ※アンダーバーも不要　※大文字
        Public propety Name As String
    End Class

    'インスタンス化
    Dim person As New Hoge
    person.Name = "ヤマダ"
    　
    '●4.1.9 Sharedフィールド（共有フィールド/Shared変数）
    'インスタンスの有無に関係なく存在、共有できる
    Public Class Hoge
        Public Shared name As String
        'Public Shared name As String = "タナカ"　　初期値ありver
    End Class

    '本来ならインスタンス化するが、その必要性がない
    'Dim person As Hoge
    'person.name = "ヤマダ" ※これは誤った使い方
    Hoge.name = "ヤマダ" 'クラスに直接アクセスできる


    '●Sharedメソッド（共有メソッド）
    'こちらもSharedフィールドと同じインスタンスの有無に関係なく存在、共有できる
    Public Class Hoge
        Public Shared Sub fnname()
            Debug.WriteLine("fugafuga")
        End Sub
    End Class

    Hoge.fname() 'クラスから直接実行！

    '●違いを整理してみる 
    Public Class Hoge
        '①フィールド
        'プロパティの作成（getとset）が要る
        'インスタンス化が要る
        Private _name As String

        '②自動実装プロパティ
        'プロパティの作成（getとset）が要らない
        'インスタンス化が要る
        '※Publicを使うこと（内部的にはPrivateのものが生成されている）
        Public Property Name As String

        '②共有フィールド（Sharedフィールド）
        'プロパティの作成（getとset）が要らない
        'インスタンス化が要らない
        Public Shared name As String
    End Class

    '●違いをSharedメソッドやインスタンスしたメソッドなどを組み合わせた実例
    Public Class Hoge
        '共有フィールド（インスタンス化なしで使える）
        Public Shared a As Integer = 100

        '共有メソッド（インスタンス化なしで使える）
        Public Shared Sub fname(n As Integer)
            a = a + n
        End Sub

        'クラスメソッド（インスタンス化しないと使えない）
        Public Sub fnameDebug()
            Debug.WriteLine(a)
        End Sub
    End Class

    Dim exInstance As New Hoge 'インスタンス化
    exInstance.fnameDebug() '100と出力
    Hoge.fname(150) '共有メソッド実行 100 + 150
    exInstance.fnameDebug() '250と出力　
    Hoge.a = 300 '共有フィールドを300に上書き
    exInstance.fnameDebug() '300と出力

    '●4.1.10 オーバーロード 
    '引数の型・数・並びが異なれば同じメソッドで定義できる
    'この「Overloads」は単なる目印なのでつけなくてもよいが可読性を踏まえつけることを推奨する
    Public Overloads Sub Fname(arg1 As Integer)
        Debug.WriteLine(arg1)
    End Sub

    Public Overloads Sub Fname(arg2 As Integer, arg3 As String)
        Debug.WriteLine(arg2 & arg3)
    End Sub

    Public Overloads Sub Fname(arg4 As Integer, arg5 As Integer)
        Debug.WriteLine(arg4 + arg5)
    End Sub

    Fname(2,3) '3番目のメソッドが実行
    Fname(12345) '1番目のメソッドが実行
    Fname(100,"円") '2番目のメソッドが実行
    Fname(200,"個数")

    '●4.2.1 コンストラクタ
    'インスタンス生成時に実行する初期化メソッドのこと
    'これだけは最初に実行しておきたい処理を記述する
    '多くの場合、パラメータを受け取ってフィールドに代入するといった使い方をされる

    Public Class SampleClass
        Public Property Name As String
        Public Property Age As Integer
    End Class

    '●基本形
    'Sub New(パラメータ)  '※ここでは修飾子省略している（Private扱い）
    '   '処理()
    'End Sub

    '●コンストラクタなし
    Public Class SampleClass
        Public Property Name As String
        Public Property Age As Integer
    End Class

    'インスタンス化
    Dim instance As New SampleClass
    instance.Name = "山田"
    instance.Age = 28


    '●コンストラクタを使った場合
    Public Class SampleClass
        '※PublicでもPrivateでもコンストラクタでアクセスできる？
        Public Name As String
        Public Age As Integer
        Private _country As String

        'コンストラクタ
        Public Sub New(arg1 As String, arg2 As Integer, arg3 As String)
            Name = arg1       'Me.Name = arg1でもOK
            Age = arg2
            _country = arg3
        End Sub
    End Class

    'インスタンス化
    Dim instance As New SampleClass("山田", 28, "Japan")

    '●フィールドの数とコンストラクタの引数（パラメータ）は違ってもよい
    Public Class SampleClass
        Public Name As String
        Public Age As Integer
        Private _country As String

        'フィールド3つに対してパラメータは1つ
        Public Sub New(arg1 As String)
            Name = arg1
        End Sub
    End Class

    'インスタンス化
    Dim instance As New SampleClass("山田") '山田だけ定義される

    '●パラメータを使わずリテラルで（この場合コンストラクタに直接）初期化することもできる
    Public Class SampleClass
        Public Property Name As String
        Public Property Age As Integer

        Public Sub New(arg1 As String)
            Name = arg1
            Age = 20
        End Sub
    End Class

    'インスタンス化
    Dim instance As New SampleClass("山田") '山田と20が定義される

    '●パラメータなしでもOK
    Public Class SampleClass
        Public Name As String
        Public Age As Integer

        'パラメータは無し
        Public Sub New()
            Name = "tanaka"
            Age = 24
        End Sub
    End Class

    'インスタンス化                                                  
    Dim instance As New SampleClass() 'tanakaと24が定義される                            

    '●コンストラクタ内に処理を追加してもよい（Return以外）
    Public Class SampleClass
        Public Name As String
        Public Age As Integer

        Public Sub New(arg1 As String, arg2 As Integer)
            Name = arg1
            Age = arg2
            MsgBox("メッセージです")  '処理を追加
        End Sub
    End Class

    '●パラメータが配列の場合
    Public Class SampleClass
        Friend names() As String
        Friend ages() As Integer

        Public Sub New(arg1() As String, arg2() As Integer)
            names = arg1
            ages = arg2
        End Sub

        Public Sub Fname()
            For Each Name In names
                Debug.WriteLine(Name)
            Next
            For Each Age In ages
                Debug.WriteLine(Age)
            Next
        End Sub
    End Class

    'インスタンス化
    Dim instance As New SampleClass({"yamada", "tanaka", "suzuki"}, {12, 41})


    '●4.2.2 コンストラクタのオーバーロード   
    Class SampleClass
        Private firstName As String
        Private lastName As String
        Private age As Integer

        '引数の異なるコンストラクタが３つある
        Public Sub New(args1 As String, args2 As String, arg3 As Integer)
            firstName = args1
            lastName = args2
            Age = arg3
        End Sub

        Public Sub New(arg1 As Integer)
            Age = arg1
        End Sub

        Public Sub New(arg1 As String, arg2 As Integer)
            firstName = arg1
            lastName = "佐藤"
            Age = arg2
        End Sub

    End Class

    Dim instance1 As New SampleClass("一郎", 23) '3番目が適応
    Dim instance2 As New SampleClass("角栄", "田中", 54) '1番目が適応
    Dim instance3 As New SampleClass(40) '2番目が適応



    '●4.3 名前空間
    '名前空間の中に名前空間を入れる場合は入れ子にするか、ドットで区切る
    Namespace SampleSpace
        Namespace Aaa
            Public Class SampleClass
                Private Shared Name As String
                Private Shared Age As Integer
                Public Sub Fname()
                    Debug.WriteLine("名前空間からのメソッド")
                End Sub
            End Class
        End Namespace
    End Namespace

    '上と同じ
    Namespace SampleSpace.Aaa
        Public Class SampleClass
            Public Property name As String
            Public Property age As Integer
            Public Sub Fname()
                Debug.WriteLine("名前空間からのメソッド")
            End Sub
        End Class
    End Namespace

    '●名前空間のクラスからインスタンス化
    '上の名前空間からインスタンスを作成した例
    Public Class Foo
        Dim instance As New SampleSpace.Aaa.SampleClass '名前空間から記述する
    instance.name = "yamada"
    instance.Fname()
    End Class


    '●4.4.1 インスタンス自体をパラメータ（引数）にすることができる
    '例）引数にインスタンス（型はクラス）にしたメソッド
    'まずはあるクラスからインスタンスを生成する
    Public Class SamplClass
        Public name As String
        Public address As String

        Public Sub New(arg1 As String, arg2 As String)
            Me.name = arg1
            Me.address = arg2
        End Sub

        Public Sub Fname()
            Debug.WriteLine("メソッドの実行")
        End Sub
    End Class

    Dim instance As New SamplClass("yamada", "東京")

    'ここでこのinstaceを引数にしたメソッドを作成する
    Public Sub Fname(arg As SamplClass) 'ここのargの方はClass型
        Debug.WriteLine(arg)
        Debug.WriteLine(arg.name) 'publicだからアクセスできることに注意
        Debug.WriteLine(arg.address)
    End Sub


    '例）フィールドがprivateでメソッドからアクセスする場合
    Public Class SamplClass
        Private name As String
        Private address As String

        Public Sub New(arg1 As String, arg2 As String)
            Me.name = arg1
            Me.address = arg2
        End Sub

        Public Sub Fname()
            Debug.WriteLine(arg1 & arg2)
        End Sub
    End Class

    Dim instance As New SamplClass("yamada", "東京")

    Public Sub Fname(arg As SamplClass)
        Debug.WriteLine(arg)
        arg.Fname() '「yamada東京」と出力　メソッドを介することでフィールドの値を出力できた
    End Sub

    '●違うクラスから生成されたインスタンス自体も引数にとって別のクラスで利用できる
    Public Class SampleClass
        Private age As Integer

        Public Sub New(arg As Integer)
            Me.Age = arg
        End Sub

        Public Sub Fname()
            Debug.WriteLine(Age)
        End Sub
    End Class

    Public Class AnotherClass
        Public Sub AnotherFname(arg As SampleClass)
            Debug.WriteLine(arg) '「一つ目のインスタンス」が出力
            arg.Fname() '一つ目のインスタンスメソッドが実行される
        End Sub
    End Class

    '一つ目のクラスからインスタンスを生成
    Dim instance As New SampleClass(55)

    '次に二つ目のクラスから別インスタンスを生成
    Dim anotherInstance As New AnotherClass
    'ここで別インスタンスのメソッドを実行するが、引数として一つ目のインスタンスをとっている
    anotherInstance.AnotherFname(instance) '形式としては[  ②インスタンス.②メソッド(①インスタンス)   ]


    '●同じクラスから生成インスタンスでも引数利用できる
    Public Class SampleClass
        Private n As Integer

        Public Sub New(arg As Integer)
            Me.n = arg
        End Sub

        Public Sub Fname()
            Debug.WriteLine(n)
        End Sub

        Public Sub AnotherFname(arg As SampleClass)
            'ここではこれらのメソッドは二つ目のインスタンスメソッドが呼ばれたときに実行される
            Debug.WriteLine(n) '「55」が出力（一つ目のインスタンス）
            Debug.WriteLine(arg.n) '「40」が出力（二つ目のインスタンス）
            Debug.WriteLine(n + arg.n) '「95」が出力
        End Sub
    End Class


    Dim instance As New SampleClass(55)

    '同じクラスから別インスタンス生成
    Dim anotherInstance As New SampleClass(40)
    '一つ目のインスタンスを引数にして同じクラスのインスタンスを引数にしている
    anotherInstance.AnotherFname(instance) '形式としては[  別インスタンス.元メソッド(元インスタンス)   ]



    '●インスタンスは代入できる
    Dim instance As New SamplClass
    Dim another As SamplClass = instance 'anotherには最初に生成したインスタンスが代入されている

    'なのでfunctionでも返り値で返せる
    Public Function Fname() As SamplClass 'この戻り値は「生成されたインスタンス」
        Dim instance As New SamplClass
        Return instance
    End Function


    '●4.4.2 インスタンスを戻り値として返す
    Public Class SampleClass
        Public n As Integer = 100

        Public Function Fname() As SampleClass
            Dim instance As New SampleClass
            Return instance
        End Function
    End Class

    'インスタンスを生成
    Dim instance As New SampleClass
    'その生成されたインスタンスFunctionプロシージャの戻り値をxxxに代入できる
    Dim xxx As SampleClass = instance.Fname()
    'Dim xxx = instance.Fname()　分かりやすくするとこういうこと（暗黙の型変換）
    'xxxの実態は「instance」

    '●4.4.3 配列の中身をクラスで構成する
    Public Class SampleClass
        Public name As String

        Public Sub New(arg As String)
            Name = arg
        End Sub
    End Class


    '3つクラスが入れる配列を宣言
    Dim arrays(2) As SampleClass

    arrays(0) = New SampleClass("tanaka")
    arrays(1) = New SampleClass("yamada")
    arrays(2) = New SampleClass("suzuki")

    'その3ループで出力
    Public Sub Fname()
        For i = 0 To 2
            Debug.WriteLine(arrays(i).Name)
        Next

        '配列個数を可変にする方法
        For i = 0 To UBound(arrays)
            Debug.WriteLine(arrays(i).Name)
        Next

        'lengthを使う方法
        For i = 0 To arrays.Length - 1
            Debug.WriteLine(arrays(i).Name)
        Next
    End Sub

    '●複数の異なるクラスも配列にできる
    Public Class SampleClass1
        Public name As String

        Public Sub New(arg As String)
            name = arg
        End Sub
    End Class

    Public Class SampleClass2
        Public age As Integer

        Public Sub New(arg As Integer)
            age = arg
        End Sub
    End Class

    Public Class SampleClass3
        Public address As String

        Public Sub New(arg As String)
            address = arg
        End Sub
    End Class

    '違うクラスをオブジェクト型の変数にいれる
    Dim objects(2) As Object
    objects(0) = New SampleClass1("yamada")
    objects(1) = New SampleClass1(23)
    objects(2) = New SampleClass1("tokyo")

    'まとめてもOK
    Dim anotherObjects() As Object = {
        New SampleClass1("tanaka"),
        New SampleClass2(48),
        New SampleClass3("osaka")
    }

    '●4.5.1 継承（サブクラス/派生クラス）
    Public Class sample
        Public Sub Fname()
            Debug.WriteLine("hoge")
        End Sub
    End Class

    '継承クラス
    Public Class SubSample
        Inherits sample

        Public Sub AnotherFname()
            Debug.WriteLine("fugafuga")
        End Sub
    End Class

    Dim instance As New SubSample
    '継承したクラスに Fname()は存在しないが継承しているので使える
    '勿論、AnotherFname()も使える
    instance.Fname()

    '●親クラスのコンストラクタを実行する(MyBase)
    Public Class sample
        Private Property name As String

        Public Sub New(arg As String)
            name = arg
        End Sub

        Public Sub Fname()
            Debug.WriteLine(name)
        End Sub
    End Class

    '継承クラス
    Public Class SubSample
        Inherits sample

        '親のコンストラクタを実行する
        Public Sub New(arg2 As String)
            MyBase.New(arg2) 'MyBaseを使用して実行する
        End Sub
    End Class

    Dim instance As New SubSample("yamada")
    instance.Fname() '「yamada」と出力 スーパークラス（親クラス/基本クラス）のメソッドを実行している


    '●サブクラスにも独自のコンストラクタがある場合
    Public Class sample
        Private Property name As String

        Public Sub New(arg As String)
            name = arg
        End Sub

        Public Sub Fname()
            Debug.WriteLine(name)
        End Sub
    End Class

    '継承クラス
    Public Class SubSample
        Inherits sample

        Private Property age As Integer
        Private Property address As String

        '引数（パラメータ）に親と子のコンストラクタをすべて用意する
        '引数は順不同だが、親→子の順序を推奨
        Public Sub New(arg1 As String, arg2 As Integer, arg3 As String)
            MyBase.New(arg1) '親クラスのコンストラクタ
            age = arg2
            address = arg3
        End Sub
    End Class

    'Dim instance As New sample("yamada", 34, "tokyo") これは不正解 スーバークラスの引数は1つだけだから
    Dim instance As New SubSample("yamada", 34, "tokyo") 'サブクラスは(1 + 2)つの引数をとる

    '●親クラスのメソッドを実行する(MyBase)
    Public Class sample
        Public Sub Fname()
            Debug.WriteLine("hoge")
        End Sub
    End Class

    '継承クラス
    Public Class SubSample
        Inherits sample

        Public Sub ChildFname()
            '親のメソッドを実行 MyBaseはつけなくても構わないが、明示する方がよい
            MyBase.Fname() 'どちらでもよい
            Fname()        'どちらでもよい
        End Sub
    End Class

    Dim instance As New ISample
    instance.ChildFname()


    '●クラスの継承を禁止する(NotInheritable)
    '禁止を宣言するのは親クラス
    Public NotInheritable Class sample
        Public Sub Fname()
            Debug.WriteLine("hoge")
        End Sub
    End Class

    '子クラスはこの時点で継承できない
    Public Class SubSample
        Inherits sample 'この時点でエラー

    End Class

    '●継承しなければ使えない(MustInherit) ※抽象クラス
    '※詳細は4.7を参照
    '継承の強制を宣言するのは親クラス
    '親クラスからインスタンスは生成できない 子クラスから生成を強制する
    Public MustInherit Class sample

    End Class

    '継承クラス
    Public Class SubSample
        Inherits sample

        Public Sub ChildFname()
            Debug.WriteLine("fuga")
        End Sub
    End Class

    Dim instance As New sample '親からインスタンス生成しようとしているのでエラー


    '●4.6.1 オーバーライド（再定義）
    '概要
    '子クラスに親クラスと同名のメソッドを作成し上書きする

    'メリット
    '親クラスとそこから派生したクラスに共通のメソッドを持たせつつカスタムできる
    '似た処理を同名で管理しつつカスタムできるので混乱しにくい

    '条件
    '親クラスとメソッドは同名
    'メソッドとパラメータは同じ構成
    '戻り値がある場合はその方が同じ
    'アクセス修飾子は変更不可

    '親クラスのメソッドに「Overrideble」を宣言し子クラスに許可する
    '子クラスのメソッドには「Overrides」を付加し上書きすることを宣言する

    Public Class sample
        Public Overridable Function Fname() As String
            Return "hoge"
        End Function
    End Class

    '継承クラス
    Public Class SubSample
        Inherits sample

        Public Overrides Function Fname() As String
            Return "fuga"
        End Function
    End Class

    Dim instance As New SubSample
    instance.Fname() '「fuga」 同じFname()でも上書きされたメソッドが実行される



    '●4.6.3 オーバーライドと親クラスのメソッド呼び出し（MyBase）
    Public Class sample
        Public Overridable Sub Fname()
            Debug.WriteLine("hoge")
        End Sub
    End Class

    '子クラス
    Public Class SubSample
        Inherits sample

        Public Overrides Sub Fname()
            Debug.WriteLine("fuga")
            MyBase.Fname() 'MyBaseで親のメソッドを実行
        End Sub
    End Class

    Dim instance As New SubSample
    instance.Fname() '「fuag hoge」 ここでは子クラスのメソッド内で親クラスのメソッドも実行


    '●オーバーロードとオーバーライドの違いを整理
    'オーバーロード：パラメータが異なることで同名メソッドで使いまわせる
    'オーバーライド：子クラスでの上書き

    '●子クラスでオーバーロード（同名メソッドの引数違い）
    Public Class sample
        Public Overloads Sub Fname(arg1 As String)
            Debug.WriteLine(arg1)
        End Sub
    End Class

    '子クラス
    Public Class SubSample
        Inherits sample

        Public Overloads Sub Fname(arg1 As String, arg2 As Integer)
            Debug.WriteLine(arg1 & arg2)
        End Sub
    End Class

    Dim instance As New SubSample '子クラスをインスタンス化
    instance.Fname("yamada") '親クラスの同名メソッドが実行（※子クラスのインスタンスでも親クラスのメソッドは実行できる）
    instance.Fname("yamada", 24) '子クラスの同名メソッドが実行

    '●4.6.5 親クラスと同名のフィールドの定義



    '●4.7 1 抽象クラス
    '役割：定義が完結していないクラス 継承先でメソッドの内容を定義を強制する
    '抽象クラスからインスタンスは生成できない
    '抽象クラスには通常のフィールドやメソッドを含めてもよい
    '抽象メソッドは必須ではない（その逆は必須 下の②参照）
    '「抽象メソッドのみ（フィールドもなし）」でもOKだが、それはインターフェイス（抽象クラスの特殊版）に分類されるのでそちらに変更すべき


    '①通常のクラスと抽象クラス違い

    '通常のクラスを継承するパターン
    Public Class sample
        Public Sub Fname()
            Debug.WriteLine("super")
        End Sub
    End Class

    Public Class SubSample
        Inherits sample

        Public Sub Fname2()
            Debug.WriteLine("sub")
        End Sub
    End Class

    '抽象クラスを継承するパターン
    Public MustInherit Class AbstractSample '抽象クラス
        '...
    End Class

    Public Class FigurativeSample '具象クラス
        Inherits AbstractSample

        Public Sub Fname()
            Debug.WriteLine("sub")
        End Sub
    End Class

    '②オーバーライドメソッドと抽象メソッドの違い
    '抽象メソッドは抽象クラス内で使うので必ずセットになる
    'なので親の抽象メソッドで「MustOverride」を使うためには親クラスで「MustInherit」を宣言が必須

    '通常のオーバーライドするパターン
    Public Class sample
        Public Overridable Sub Fname()
            Debug.WriteLine("hoge")
        End Sub
    End Class

    Public Class SubSample
        Inherits sample

        Public Overrides Sub Fname()
            Debug.WriteLine("fuga")
        End Sub
    End Class

    Dim instance As New SubSample
    instance.Fname() '子クラスメソッドは親メソッドを上書きしている

    '抽象クラスと抽象メソッドを使ったパターン
    Public MustInherit Class AbstractSample 'MustInheritを付加
        '親メソッドにMustOverrideを付加
        'メソッドの定義は未完成のまま（空のメソッド）で、継承した子クラス（具象クラス）で定義を強制する
        Public MustOverride Sub Fname()
    End Class

    Public Class FigurativeSample
        Inherits SubSample

        Public Overrides Sub Fname() 'このオーバーライドをメソッドの「実装」と呼ぶ
            Debug.WriteLine("sub")
        End Sub
    End Class

    Dim instance As New AbstractSample '抽象クラスからインスタンスは生成できないのでエラー
    Dim instance As New FigurativeSample '具象クラスなのでOK



    '●4.7.2 インターフェイス
    '抽象クラスの特殊版

    '違い
    '抽象クラス：抽象メンバ以外（通常のフィールドやメソッド）も含めることができる
    'インターフェイス：抽象メソッドだけを持つ

    'ポイント
    '抽象メソッドにアクセス修飾子はつけられない
    '慣習的に「I」から始まる名前にする
    '親クラスに「Interface」、子クラスと子メソッドに「Implements」を付加
    '構成がやや異なる（※に注意）
    Public Interface ISample '※親クラスにInterfaceを付加 Class不要
        '※publicもMustOverrideも付加できない
        Sub Fname(n As Integer)
    End Interface

    Public Class SubSample
        Implements ISample '※子クラスにImplementsを付加

        Public Sub Fname(n As Integer) Implements ISample.Fname '※子メソッドにImplementsと「親クラス.親メソッド」を付加
            Debug.WriteLine(n)
        End Sub
    End Class

    Dim instance As New SubSample
    instance.Fname(10)


    '●aaa

    '●


    '●


    '●


    '●


    '●

End Namespace














