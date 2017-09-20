# Supervised-Learning

## Single layer feed forward neural network

Requires: [cANN.cls](Modules/cANN.cls), [modMath.bas](../../../Scientific-Toolkit/blob/master/Modules/modMath.bas)

![ANN1](Screenshots/ANN1.jpg)

Test data here is wine data set from [UCI Machine Learning Repository](https://archive.ics.uci.edu/ml/datasets.html)<sup>1</sup>. It consists of 178 samples of wines made from three different cultivars, which will be named as classes 1, 2 and 3 in the following sections. 13 attributes of these wine samples were measured. Data is read in as an array of *x(1:N,1:D)* where N=178 and D=13 are the number of samples and dimensions respectivly. *x_class(1:N)* holds the class label of each sample.

1. Forina, M. et al. [UCI Machine Learning Repository](http://archive.ics.uci.edu/ml). Institute of Pharmaceutical and Food Analysis and Technologies. 

### 1. Prep the data

Before we start, first normalize data x() and use `modmath.Class2Vec` to convert the class labels to vectors of 0 and 1. So class 1 is represented by a vector of (1,0,0), class 2 by (0,1,0) and class 3 by (0,0,1). The vector is stored in *x_class_vec()'.

```
  Call modMath.Normalize_x(x, x_shift, x_scale, "AVGSD") 'Normalize data to zero mean and unit variance
  Call modMath.Class2Vec(x_class, x_class_vec, n_output, class_map) 'vectorize labels
```

### 2. Split the data into training, validation and test set

Remember that the data you get may have been sorted in some way. In this case the data from UCI is sorted by class label. If you directly split the data in half you will end up with samples bias to a certain class. So it's a good ida to shuffle the dataset first.

```
    iTrain = modMath.index_array(1, n_raw)  'generate pointers 1 to N for each data
    Call modMath.Shuffle(iTrain) 'Shuffle the pointers
```
Now split the set into training/validation/test set by portion of 70/30/78.

```
    Call modMath.MidArray(iTrain, 71, 100, iValid)  'pointer to validation set
    Call modMath.MidArray(iTrain, 101, 178, iTest)  'pointer to test set
    ReDim Preserve iTrain(1 To n_train)             'pointer to training set
    
    'Separate data into train, validation and test set
    Call modMath.Filter_Array(x, x_train, iTrain)
    Call modMath.Filter_Array(x_class, x_class_train, iTrain)
    Call modMath.Filter_Array(x_class_vec, x_class_vec_train, iTrain)
    
    '...repeat for validation and test set
```

### 3. Train neural network
Now we can feed the training and validation set into our ANN. We will set the number of hidden units to 13 in this example. Figure on the above left shows the architecture of this network. Activation functions are hard coded to be sigmoid and softmax in the hidden and output layer respectively.

Syntax to initialize and train the network is as below:
```  
  Dim ANN1 As New cANN
  With ANN1
    Call .Init(n_input, n_output, 13)     'Initialize
    Call .Trainer(x_train, x_class_vec_train, , , , , , x_valid, x_class_vec_valid)
    cost_function=.cost_function
  End With
```

The cost function at every epoch can is pulled out and shown on the upper right figure. Note how the cost function continued to drop in the training set after about 600 epochs but stop dropping in the validation set. This could be a sign of overfitting, which is why we want to have a validation set to make sure we know when to stop training.

Now we are ready to test the model on the training set
```      
  Call ANN1.InOut(x_train, y)
  Call modMath.Vec2Class(y, class_map, x_class_out) 'Recover class label
```
The two charts below show the accuracy of our trained network. Accuracy on traing set is 100%. On the test set, it makes one incorrect prediction, it misclassifies a class2 sample as class3.

![ANN2](Screenshots/ANN2.jpg)

### 4. Save trained network
The trained network weights can be print to an Excel worksheet with
```
Call ANN1.Print_Model(wksht)
```

which can be reused next time by the read command
```
Call ANN1.Read_Model(wksht)
```

