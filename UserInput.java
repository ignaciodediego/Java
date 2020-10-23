public class UserInput {
    
	public String value = "";
	
    public static class TextInput extends UserInput{
        
        public void add(char c){
            
            this.value += Character.toString(c);
        }
        public String getValue(){
            
            return this.value;
        }
    }

    public static class NumericInput extends TextInput{
        
        public void add(char c){
            
            if(Character.isDigit(c)){
                this.value += c;
            }
        }       
    }

    public static void main(String[] args) {
        TextInput input = new NumericInput();
        input.add('1');
        input.add('a');
        input.add('0');
        System.out.println(input.getValue());
    }
}