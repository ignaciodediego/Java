public class Song {
    private String name;
    private Song nextSong;

    public Song(String name) {
        this.name = name;
    }

    public void setNextSong(Song nextSong) {
        this.nextSong = nextSong;
    }

    public boolean hasNext() {
        return nextSong != null;
    }

    public boolean isRepeatingPlaylist() {
        // It's called Floyd's Cycle-Finding Algorithm, but it's sometimes referred to as "The Tortoise and the Hare Algorithm".
        Song fast = this;
        Song slow = this;

        while (true) {
            if (slow.hasNext()) slow = slow.nextSong; else return false;

            if (fast.hasNext()) fast = fast.nextSong; else return false;
            if (fast.hasNext()) fast = fast.nextSong; else return false;

            if (slow == fast) return true;
        }
    }

    public static void main(String[] args) {
        Song first = new Song("Hello");
        Song second = new Song("Eye of the tiger");

        first.setNextSong(second);
        second.setNextSong(first);

        System.out.println(first.isRepeatingPlaylist());
    }
}