import numpy as np

# Write a function that takes as input two lists Y, P,
# and returns the float corresponding to their cross-entropy.


def binary_cross_entropy(Y, P):
    Y = np.float_(Y)
    P = np.float_(P)
    return (-np.sum((Y * np.log(P)) + ((1 - Y)*np.log(1 - P))))


def multiclass_cross_entropy(Y, P):
    Y = np.float_(Y)
    P = np.float_(P)
    return (-np.sum(np.sum((Y * np.log(P)))))

# Cross entropy tending to zero is good measure of classfiication
# Cross entropy tending to infinity is bad measure of classfiication


# In bellow example, data point 1,3,4 are said to be correctly classfiied [classification truth is 1]... yet their probabilities are not high
# While, data point 2 is said to be misclassified yet its probability is high
# Hence, CE measure returned is (4.828313737302301) which tends towards infinity, denoting bad classification
Y = [1, 0, 1, 1]
P = [0.4, 0.6, 0.1, 0.5]
CE = binary_cross_entropy(Y, P)
print("Binary CE for 2 class problem: ",CE)


# Multiclass cross entropy function inputs classification truth and probability for all datapoints for all classes  
Y = [[1, 0, 1, 1],\
    [0, 1, 0, 0]]
P = [[0.4, 0.6, 0.1, 0.5],\
    [0.6, 0.4, 0.9, 0.5]]
CE = multiclass_cross_entropy(Y,P)
print("Multiclass CE for 2 class problem",CE)
