import java.util.Stack;

public class Path {
    private String path;

    public Path(String path) {
        this.path = path;
    }

    public String getPath() {
        return path;
    }

    public void cd(String newPath) {
        if (newPath.startsWith("/") && !newPath.contains("..")) {
            path = newPath;
            return;
        }

        String[] currentPathArray = path.split("/");
        String[] newPathArray = newPath.split("/");

        Stack<String> stack = new Stack<>();

        for (String el :
                currentPathArray) {
            if (!el.isEmpty())
                stack.push(el);
        }

        for (String el : newPathArray) {
            if (el.isEmpty())
                continue;
            if (el.equals("..")) {
                stack.pop();
            } else {
                stack.push(el);
            }
        }

        StringBuilder result = new StringBuilder();
        for (String el :
                stack) {
            result.append("/").append(el);
        }

        path = result.toString();
    }

    public static void main(String[] args) {
        Path path = new Path("/a/b/c/d");
        path.cd("x");
        System.out.println(path.getPath());
    }
}