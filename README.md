# Supervised-Learning

## Single layer feedforward neural network

Requires: [cANN.cls](Modules/cANN.cls), [modMath.bas](../../../Scientific-Toolkit/blob/master/Modules/modMath.bas)

![ANN1](Screenshots/ANN1.jpg)

Test data here is wine data set from [UCI Machine Learning Repository](https://archive.ics.uci.edu/ml/datasets.html)<sup>1</sup>. It consists of 178 samples of wines made from three different cultivars, which will be named as classes 1, 2 and 3 in the following sections. 13 attributes of these wine samples were measured. Data is read in as an array of *x(1:N,1:D)* where N=178 and D=13 are the number of samples and dimensions respectivly. *x_class(1:N)* holds the class label of each sample.

1. Forina, M. et al. [UCI Machine Learning Repository](http://archive.ics.uci.edu/ml). Institute of Pharmaceutical and Food Analysis and Technologies. 

#### 1. Prep the data

Before we start, first normalize data *x()* and use `modmath.Class2Vec` to convert the class labels to vectors of 0 and 1. So class 1 is represented by a vector of (1 0 0), class 2 by (0 1 0) and class 3 by (0 0 1). The vector is stored in *x_class_vec()'.

```
  Call modMath.Normalize_x(x, x_shift, x_scale, "AVGSD") 'Normalize data to zero mean and unit variance
  Call modMath.Class2Vec(x_class, x_class_vec, n_output, class_map) 'vectorize labels
```

#### 2. Split the data into training, validation and test set

Remember that the data you get may have been sorted in some way. In this case the data from UCI is sorted by class label. If you directly split the data in half you will end up with samples bias to a certain class. So it's a good idea to shuffle the dataset first.

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

#### 3. Train neural network
Now we can feed the training and validation sets into our ANN. We will set the number of hidden units to 13 in this example. Figure on the above left shows the architecture of this network. Activation functions are hard coded to be sigmoid and softmax in the hidden and output layer respectively.

Syntax to initialize and train the network is as below:
```  
  Dim ANN1 As New cANN
  With ANN1
    Call .Init(n_input, n_output, 13)     'Initialize
    Call .Trainer(x_train, x_class_vec_train, , , , , , x_valid, x_class_vec_valid)
    cost_function=.cost_function
  End With
```

The cost function at every epoch can is pulled out and shown on the upper right figure. Note how the cost function continued to drop in the training set after about 600 epochs but stopped improving in the validation set. This could be a sign of overfitting, which is why we want to have a validation set to make sure we know when to stop training.

Now we are ready to test the model on the training set
```      
  Call ANN1.InOut(x_train, y)
  Call modMath.Vec2Class(y, class_map, x_class_out) 'Recover class label
```
The two charts below show the accuracy of our trained network. Accuracy on traing set is 100%. On the test set, it makes one incorrect prediction, it misclassifies a class2 sample as class3.

![ANN2](Screenshots/ANN2.jpg)

#### 4. Save trained network
The trained network weights can be printed to an Excel worksheet with
```
Call ANN1.Print_Model(wksht)
```

which can be reused next time by the read command
```
Call ANN1.Read_Model(wksht)
```

## Recurrent Neutral Network with Long-Short-Term-Memory (LSTM) unit

For this example we experiment with natural language processing (NLP), using text from "Alice in Wonderland". To keep things simple, only the first few paragraphs are used. That contains ~4000 characters including punctuations and spaces. All letters are converted to lower case. Linebreaks are removed.

The stream of charaters are stored in a vector *strArr(1:N)*, which is then converted to a binary vector with a codebook, using:
```
Call modMath.Class2Vec(strArr, strVec, n_dimension, codebook)
```
So *strVec(1:N, 1:n_dimension)* is an array of 0 and 1. Where the non-zero position corresponds to a character in the code book, which is a vector of length n_dimenison.

The full data is then split into segments for batch training. In this example let's say we use sequences of length 30 and the goal is to predict the 31-st character. So we split the full series into set of sequences of length 30, and another set also of length 30 but offset by 1 step as the target sequence.
```
Call Sequence2Segments(strVec, y_input, 30, 1, 3)   'segments of length 30 starting from the 1st position
Call Sequence2Segments(strVec, y_target, 30, 2, 3)  'segments of length 30 starting from the 2nd position
```

Now we are ready to train a LSTM network. There are two classes of LSTM implemeted: cLSTM.cls and cLSTM_B.cls for unidirectional and bidirectional network. The loss function used is mutliclass cross entropy with a softmax output layer. The implementation is fairly basic without any dropout or regularization. Let's just see what that gets to:
```
Dim LSTM1 As cLSTM
Set LSTM1 = New cLSTM
With LSTM1
    Call .Init(n_dimension, 80, n_dimension)    'LSTM with 80 hidden units
    Call .Train_Batch(xS, y_tgt, , 0.01, , 20)  'train for 20 epochs with learning rate of 0.01
    cost_function = .cost_function              'Print model
    .Print_Model(wksht)
End With
```
After the training is done. Use `.cost_function` to print the cost function to see that it has converged. Also use `.Print_Model(wksht as WorkSheet)` to print the trained weights and save it. This is important since you certainly don't want to spend another day to train it from scratch. With a saved model, you can read it back it with `.Read_Model(wksht)`, then use it or continue to train it. Let's just try it to see what it's learnt.

When the network was given the keywords "alice" and "rabbit" and asked to generate a sentences of 50 characters using the command 'LSTM1.Generate(strSeed, 50, 20)', it generates this: 

|Input | Output |
|------|--------|
| alice | alice to herself, \`at thought the ent out again. t|
| rabbit | rabbit-hole ont mooked at the tith the antipat thi|

Well, not very meaningful...but at least it gets some spellings correct. So now you have it : your own chatbot that talks garbage! My wife used to talk to a Christian Grey chatbot on Facebook which sounds even more garb-lish. At least with this one you can train it on your favorite topics.

Obviously a real deal NLP will need much more training sample, the training model will be more invovled, and you definitely will not run a language training model in VBA. I just want to show you how it works.

In fact if you are working in the financial industry, you are more likely to be predicting continuous signal (i.e. price) instead of categorical signal. In that case the model can be easily modified to use a mean-sqaure-error loss function and a sigmoid or linear output activation.

## Commonly used loss function and activation function at the output layer
Here are some commonly used functions and their respective derivatives and deltas listed for easy reference. *t<sub>i</sub>* is the target output at node *i*, *y<sub>i</sub>* and *x<sub>i</sub>* are the output and input to node *i*.

![Eq1](Screenshots/BackProp_Eq1.jpg)

To calculate gradient of the loss function with respect to weights at the *k*-th layer, we backpropagate the gradient using:

![Eq2](Screenshots/BackProp_Eq2.jpg)

To put in words, gradient of weight *w<sub>ji</sub>* is equal to its input multiplied by the delta at the exiting node. And delta at node *i* of the *k*-th layer is given by weighted sum of deltas from connected nodes in the next layer, modulated by its own gradient.
