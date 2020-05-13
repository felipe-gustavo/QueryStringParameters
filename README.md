# QueryStringParameters

Work with Query String Parameters in VBA easily

## Dependencies

Microsoft Scripting Runtime


## How to use

After import the file (download [here](https://github.com/felipe-gustavo/QueryStringParameters/archive/master.zip)) to your VBA project, let's go to the examples.

Create a Object:
```vb
Dim objQueryStr as QueryStringParameters

Set objQueryStr = New QueryStringParameters
```

#### Importing a query string to object
```vb 
Dim objQueryStr as QueryStringParameters

Set objQueryStr = New QueryStringParameters

objQueryStr.parseQueryStringParameters "foo[bar]=value1&foo[menu]=value2"
```
The below examples has based in above command 


#### Getting value in string
```vb
Debug.Print objQueryStr.toString()
' >output: foo[bar]=value1&foo[menu]=value2
```

#### Adding new field in query string
```vb
objQueryStr.add "myValue", Array("foo", "top")

Debug.Print objQueryStr.toString()
' >output: foo[bar]=value1&foo[menu]=value2&foo[top]=myValue
```

#### Adding new field in query string as first field
```vb
objQueryStr.add "mySecondValue", Array("foo", "bottom"), objQueryStr.addAsFirstValue

Debug.Print objQueryStr.toString()
' >output: foo[bottom]=mySecondValue&foo[bar]=value1&foo[menu]=value2&foo[top]=myValue
```

#### Adding new field in query string after a field
```vb
objQueryStr.add "mySecondValue2", Array("foo", "bottom2"), "bottom"

Debug.Print objQueryStr.toString()
' >output: foo[bottom]=mySecondValue&foo[bottom2]=mySecondValue2&foo[bar]=value1&foo[menu]=value2&foo[top]=myValue
```
Note: If the field doesn't exist, the new value will be inserted in last field

#### Adding new sequential field
```vb

Set objQueryStr = New QueryStringParameters

objQueryStr.add "value1", Array("foo", "")
objQueryStr.add "value2", Array("foo", "")

Debug.Print objQueryStr.toString()
' >output: foo[]=value1&foo[]=value2
```

#### Adding new value after sequential field
```vb
objQueryStr.add "newValue", Array("foo", "new"), .getSequentialKeyByIndex(0, array("foo"))

Debug.Print objQueryStr.toString()
' >output: foo[]=value1&foo[new]=newValue&foo[]=value2
```

#### Removing fields
```vb
objQueryStr.remove Array("foo", "new")
objQueryStr.remove Array("foo", .getSequentialKeyByIndex(0, array("foo")))

Debug.Print objQueryStr.toString()
' >output: foo[]=value1
```


#### Updating fields
```vb
objQueryStr.update 2, Array("foo", "menu")                                      '' Auto add
objQueryStr.update 1, Array("foo", .getSequentialKeyByIndex(0, array("foo")))   '' Update sequential key

Debug.Print objQueryStr.toString()
' >output: foo[]=1&foo[menu]=2
```


#### Getting value
```vb
Dim objDictionary as Dictionary
Set objDictionary = objQueryStr.getValue Array("foo")

For Each Item in objDictionary
  Debug.Print Item
Next
' >output: 1
' >output: 2
```

#### Getting value nonexistent
```vb
value = objQueryStr.getValue Array("foo", "new")

Debug.Print IsError(Value)
' >output: True
```
