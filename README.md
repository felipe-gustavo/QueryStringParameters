# QueryStringParameters

Work with Query String Parameters in VBA easily

## How to use

After import the file ([here](https://raw.githubusercontent.com/felipe-gustavo/QueryStringParameters/master/QueryStringParameters.cls)) to you VBA project, let's go to the examples.

Create a Object:
```vb
Dim objQueryStr as QueryStringParameters

Set objQueryStr = New QueryStringParameters
```

#### Importing a query string to object:
```vb
Dim objQueryStr as QueryStringParameters

Set objQueryStr = New QueryStringParameters

objQueryStr.parseQueryStringParameters "foo[bar]=value1&foo[menu]=value2"
```

#### Getting value in string:
```vb
Debug.Print objQueryStr.toString()
' >output: foo[bar]=value1&foo[menu]=value2
```

#### Adding new field in query string:
```vb
objQueryStr.add "myValue", Array("foo", "top")

Debug.Print objQueryStr.toString()
' >output: foo[bar]=value1&foo[menu]=value2&foo[top]=value2
```

#### Adding new field in query string as first field:
```vb
objQueryStr.add "mySecondValue", Array("foo", "bottom"), objQueryStr.addAsFirstValue

Debug.Print objQueryStr.toString()
' >output: foo[bar]=value1&foo[menu]=value2&foo[top]=value2
```

#### Adding new field in query string after a field:
```vb
objQueryStr.add "mySecondValue", Array("foo", "bottom"), objQueryStr.addAsFirstValue

Debug.Print objQueryStr.toString()
' >output: foo[bar]=value1&foo[menu]=value2&foo[top]=value2
```
Note: If the field doesn't exist, the new value will be inserted in last field

#### Adding new sequential field
```vb
objQueryStr.add "mySequentialValue", Array("foo", "")
objQueryStr.add "mySequentialValue2", Array("foo", "")

Debug.Print objQueryStr.toString()
' >output: foo[bar]=value1&foo[menu]=value2&foo[top]=value2&foo[]=mySequentialValue&foo[]=mySequentialValue2
```

#### Adding new value after sequential field
```vb
objQueryStr.add "newValue", Array("foo", "new"), .getSequentialKeyByIndex(0, array("foo"))

Debug.Print objQueryStr.toString()
' >output: foo[bar]=value1&foo[menu]=value2&foo[top]=value2&foo[]=mySequentialValue&foo[new]=newValue&foo[]=mySequentialValue2
```
