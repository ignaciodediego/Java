public class MergeNames {
	
	public static String[] uniqueNames(String[] names1, String[] names2) {
        //throw new UnsupportedOperationException("Waiting to be implemented.");
		
		//String[] res = new String[names1.length + names2.length];
		
		ArrayList<String> ar = new ArrayList<String>();

        //System.arraycopy( names1, 0, res, 0, names1.length );
		//System.arraycopy( names1, 0, ar, 0, names1.length );
        //System.arraycopy( names2, 0, res, names1.length, names2.length );
		
		for (int a = 0; a < names1.length; ++a) {

				ar.add(names1[a]);
		}
		
		for (int i = 0; i < names2.length; ++i) {

			if(!Existe(names2[i], ar)) {

				ar.add(names2[i]);
			}
		}
        return ar.toArray(new String[ar.size()]);
    }
    
	public static Boolean Existe(String valor, ArrayList<String> listaElem) {
		
		
		Boolean encontrado = false;
		
		for (int i = 0; i < listaElem.size(); ++i) {

			if(listaElem.get(i).equals(valor)) {

				encontrado = true;
				break;
			}
			
		}
		
		return encontrado;
	}
	
    public static void main(String[] args) {
        String[] names1 = new String[] {"Ava", "Emma", "Olivia"};
        String[] names2 = new String[] {"Olivia", "Sophia", "Emma"};
        System.out.println(String.join(", ", MergeNames.uniqueNames(names1, names2))); // should print Ava, Emma, Olivia, Sophia
    }

}