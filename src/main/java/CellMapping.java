import java.util.ArrayList;
import java.util.HashMap;

public class CellMapping {
    private String column;
    private Integer row;
    private String dataColumn;
    private Boolean replaceStuff;

    public static HashMap<Integer, ArrayList<CellMapping>> mapping = new HashMap<>();

    static {
        ArrayList<CellMapping> sheetOneMapping = new ArrayList<CellMapping>() {{
            add(new CellMapping("V", 9, "N", true));
            add(new CellMapping("FM",9, "N", true));
            add(new CellMapping("FM",6, "L", false));
            add(new CellMapping("AS",18, "H", false));
            add(new CellMapping("BU",18, "E", false));
            add(new CellMapping("ED",44, "P", false));
            add(new CellMapping("FI",44, "Q", false));
            add(new CellMapping("BR",42, "V", false));
        }};
        mapping.put(1, sheetOneMapping);
        ArrayList<CellMapping> sheetTwoMapping = new ArrayList<CellMapping>(){{
            add(new CellMapping("N", 4, "N", true));
            add(new CellMapping("CO", 4, "J", false));
            add(new CellMapping("EL", 4, "K", false));
            add(new CellMapping("Q", 14, "S", true));
            add(new CellMapping("FD", 34, "W", false));
            add(new CellMapping("FD", 36, "X", false));
        }};
        mapping.put(2, sheetTwoMapping);
    }

    static HashMap<String, HashMap<String, String>> replaceMapping = new HashMap<String, HashMap<String, String >>(){{
        put("V9", new HashMap<String, String>(){{
            put("ООО \"БауИнвест\"", "ООО \"БАУИнвест\" Москва, Новинский бульвар, дом 31, эт. 8, пом. I, комн. 41 телефон: +7 (495) 782-17-44");
            put("ООО \"Автомиг\"", "TEST REPLACE");
        }});
        put("FM9", new HashMap<String, String>(){{
            put("ООО \"БауИнвест\"", "37189067");
            put("ООО \"Автомиг\"", "TEST REPLACE");

        }});
        put("N4", new HashMap<String, String>(){{
            put("ООО \"БауИнвест\"", "ООО \"БАУИнвест\" Москва, Новинский бульвар, дом 31, эт. 8, пом. I, комн. 41 телефон: +7 (495) 782-17-44, р/с 40702810301400013896 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593");
            put("ООО \"Автомиг\"", "TEST REPLACE");

        }});
        put("Q14", new HashMap<String, String>(){{
            put("Асиньино-1", "Карьер Аксиньино-1, деревня Аксиньино, городской округ Щёлково, Московская область");
        }});
    }};

    public CellMapping(String column, Integer row, String dataColumn, Boolean replaceStuff) {
        this.column = column;
        this.row = row - 1;
        this.dataColumn = dataColumn;
        this.replaceStuff = replaceStuff;
    }

    public String getColumn() {
        return column;
    }

    public Integer getRow() {
        return row;
    }

    public String getDataColumn() {
        return dataColumn;
    }

    public Boolean getReplaceStuff() {
        return replaceStuff;
    }

    public static String replaceStuff(CellMapping mapping, String value) {
        String cellId = mapping.getColumn() + (mapping.getRow() + 1);
        if (replaceMapping.containsKey(cellId)) {
            if (replaceMapping.get(cellId).containsKey(value)) {
                return replaceMapping.get(cellId).get(value);
            }
        }
        return "UNKNOWN_RESULT";
    }
}
