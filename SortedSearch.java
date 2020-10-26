import java.util.Arrays;

/*
11. Sorted Search
Implement function countNumbers that accepts a sorted array of unique integers and, efficiently with respect to time used,
counts the number of array elements that are less than the parameter lessThan.
For example, SortedSearch.countNumbers(new int[] { 1, 3, 5, 7 }, 4) should return 2 because there are two array elements less than 4.
 */
public class SortedSearch {
    private static int countNumbers(int[] sortedArray, int lessThan, int from, int to) {
        int midIndex = from + (to - from) / 2;

        if (lessThan == sortedArray[midIndex]) {
            return midIndex;
        }

        if (midIndex <= from) {
            return from + 1;
        }

        if (midIndex >= to) {
            return to + 1;
        }

        if (lessThan < sortedArray[midIndex]) {
            return countNumbers(sortedArray, lessThan, from, midIndex);
        } else {
            return countNumbers(sortedArray, lessThan, midIndex, to);
        }
    }

    public static int countNumbers(int[] sortedArray, int lessThan) {
        if (sortedArray.length <= 1 || lessThan < sortedArray[0]) {
            return 0;
        }

        if (lessThan > sortedArray[sortedArray.length-1]) {
            return sortedArray.length;
        }

        return countNumbers(sortedArray, lessThan, 0, sortedArray.length-1);
    }

    public static void main(String[] args) {
        System.out.println(SortedSearch.countNumbers(new int[] { 1, 3, 5, 7 }, 6));
    }
}