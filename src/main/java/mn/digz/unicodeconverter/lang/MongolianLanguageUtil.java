package mn.digz.unicodeconverter.lang;

/**
 *
 * @author MethoD
 */
public class MongolianLanguageUtil {
    public static String toUnicode(String input) {
        String output = "";
        try {
            if(input != null && input.length() > 0) {
                for (int i = 0; i < input.length(); i++) {
                    output += toUnicode(input.charAt(i));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return output;
    }
    
    private static char toUnicode(char input) {
        if(input == 168) {
            return 1025;
        } else if(input == 170) {
            return 1256;
        } else if(input == 175) {
            return 1198;
        } else if(input == 184) {
            return 1105;
        } else if(input == 186) {
            return 1257;
        } else if(input == 191) {
            return 1199;
        }/* else if(input == 1028) {
            return 1256;
        } else if(input == 1031) {
            return 1198;
        } else if(input == 1108) {
            return 1257;
        } else if(input == 1111) {
            return 1199;
        } */else if(192 <= input && input <= 255) {
            return (char)(input + 848);
        } else {
            return input;
        }
    }
    
    public static void main(String[] args) {
        System.out.println(toUnicode("ÀÁÂÃÄÅ¨ÆÇÈÉÊËÌÍÎªÏÐÑÒÓ¯ÔÕÖ×ØÙÚÛÜÝÞßàáâãäå¸æçèéêëìíîºïðñòó¿ôõö÷øùúûüýþÿ"));
    }
}
