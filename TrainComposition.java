public class TrainComposition {

    class Wagon {
        Wagon left;
        Wagon right;
        int id;

        Wagon(int wagonId) {
            id = wagonId;
        }
    }

    private Wagon mostLeft;
    private Wagon mostRight;

    public void attachWagonFromLeft(int wagonId) {
        if (mostLeft == null) {
            mostLeft = new Wagon(wagonId);

            if (mostRight == null) {
                mostRight = mostLeft;
            }
        } else {
            Wagon oldMostLeft = mostLeft;
            mostLeft = new Wagon(wagonId);
            oldMostLeft.left = mostLeft;
            mostLeft.right = oldMostLeft;
    }

        System.out.println("Most left="+mostLeft.id+" most right="+mostRight.id);
    }

    public void attachWagonFromRight(int wagonId) {
        if (mostRight == null) {
            mostRight = new Wagon(wagonId);

            if (mostLeft == null) {
                mostLeft = mostRight;
            }
        } else {
            Wagon oldMostRight = mostRight;
            mostRight = new Wagon(wagonId);
            oldMostRight.right = mostRight;
            mostRight.left = oldMostRight;
        }
        System.out.println("Most left="+mostLeft.id+" most right="+mostRight.id);
    }

    public int detachWagonFromLeft() {
        int detachedId = mostLeft.id;
        if (mostLeft.right == null) {
            mostRight = null;
        }
        mostLeft = mostLeft.right;
        System.out.println("Most left="+mostLeft.id+" most right="+mostRight.id);
        return detachedId;
    }

    public int detachWagonFromRight() {
        int detachedId = mostRight.id;
        if (mostRight.left == null) {
            mostLeft = null;
        }
        mostRight = mostRight.left;
        System.out.println("Most left="+mostLeft.id+" most right="+mostRight.id);
        return detachedId;
    }

    public static void main(String[] args) {
        TrainComposition tree = new TrainComposition();
        tree.attachWagonFromLeft(7);
        tree.attachWagonFromLeft(13);
        System.out.println(tree.detachWagonFromRight()); // 7
        System.out.println(tree.detachWagonFromLeft()); // 13
    }
}