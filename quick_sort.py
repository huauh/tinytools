def quick_sort(array):
    less = []
    greater = []
    if len(array) <= 1:
        return array
    pivot = array.pop()
    for x in array:
        if x <= pivot:
            less.append(x)
        else:
            greater.append(x)
    return quick_sort(less) + [pivot] + quick_sort(greater)


# Practice
def main():
    array = [45, 3, 73, 34, 7, 2, 90, 456, 23]
    print(quick_sort(array))


if __name__ == "__main__":
    main()