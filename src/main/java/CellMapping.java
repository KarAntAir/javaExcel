import java.util.ArrayList;
import java.util.HashMap;

public class CellMapping {
    private String column;
    private Integer row;
    private String dataColumn;
    private CellAction action;

    public static HashMap<Integer, ArrayList<CellMapping>> mapping = new HashMap<>();

    static {
        ArrayList<CellMapping> sheetOneMapping = new ArrayList<CellMapping>() {{
            add(new CellMapping("V", 9, "N", CellAction.REPLACE));
            add(new CellMapping("FM",9, "N", CellAction.REPLACE));
            add(new CellMapping("FM",6, "L", null));
            add(new CellMapping("AS",18, "H", null));
            add(new CellMapping("BU",18, "E", null));
            add(new CellMapping("FM",7, "M", CellAction.DAY));
            add(new CellMapping("FS",7, "M", CellAction.MONTH));
            add(new CellMapping("FZ",7, "M", CellAction.YEAR));
            add(new CellMapping("BE",44, "M", CellAction.DAY));
            add(new CellMapping("BM",44, "M", CellAction.IN_WORDS_MONTH));
            add(new CellMapping("CG",44, "M", CellAction.YEAR));
            add(new CellMapping("DS",28, "H", null));
            add(new CellMapping("CR",28, "H", CellAction.IN_WORDS));
            add(new CellMapping("FI",36, "V", null));
            add(new CellMapping("FI",44, "Q", null));
        }};
        mapping.put(1, sheetOneMapping);
        ArrayList<CellMapping> sheetTwoMapping = new ArrayList<CellMapping>(){{
            add(new CellMapping("N", 4, "N", CellAction.REPLACE));
            add(new CellMapping("CO", 4, "J", null));
            add(new CellMapping("EL", 4, "K", null));
            add(new CellMapping("Q", 14, "S", CellAction.REPLACE));
            add(new CellMapping("FD", 34, "W", null));
            add(new CellMapping("FD", 36, "X", null));
            add(new CellMapping("X", 3, "M", CellAction.DAY));
            add(new CellMapping("AD", 3, "M", CellAction.IN_WORDS_MONTH));
            add(new CellMapping("AW", 3, "M", CellAction.YEAR));
            add(new CellMapping("FP", 2, "L", null));
            add(new CellMapping("D", 23, "E", null));
            add(new CellMapping("FU", 23, "H", null));
            add(new CellMapping("FU", 26, "H", null));
            add(new CellMapping("Q", 29, "H", CellAction.IN_WORDS));
            add(new CellMapping("CF", 29, "H", CellAction.IN_WORDS));
            add(new CellMapping("DC", 30, "V", null));
            add(new CellMapping("CZ", 35, "Q", null));
        }};
        mapping.put(2, sheetTwoMapping);
    }

    static HashMap<String, HashMap<String, String>> replaceMapping = new HashMap<String, HashMap<String, String >>(){{
        put("V9", new HashMap<String, String>(){{
            put("ООО \"БауИнвест\"", "ООО \"БАУИнвест\" Москва, Новинский бульвар, дом 31, эт. 8, пом. I, комн. 41 телефон: +7 (495) 782-17-44");
            put("ООО \"Корона Рус\"", "ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7\u00A0(925)\u00A0641-85-57");
        }});
        put("FM9", new HashMap<String, String>(){{
            put("ООО \"БауИнвест\"", "37189067");
            put("ООО \"Корона Рус\"", "18703067");

        }});
        put("N4", new HashMap<String, String>(){{
            put("ООО \"БауИнвест\"", "ООО \"БАУИнвест\" Москва, Новинский бульвар, дом 31, эт. 8, пом. I, комн. 41 телефон: +7 (495) 782-17-44, р/с 40702810301400013896 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593");
            put("ООО \"Корона Рус\"", "ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7\u00A0(925)\u00A0641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593");

        }});
        put("Q14", new HashMap<String, String>(){{
            put("Асиньино-1", "Карьер Аксиньино-1, деревня Аксиньино, городской округ Щёлково, Московская область");
            put("Симбухово", "Карьер Симбухово, деревня Симбухово, Наро-Фоминский городской округ, Московская область");
        }});
    }};

    public CellMapping(String column, Integer row, String dataColumn, CellAction action) {
        this.column = column;
        this.row = row;
        this.dataColumn = dataColumn;
        this.action = action;
    }

    public String getColumn() {
        return column;
    }

    public Integer getIndexRow() {
        return row - 1;
    }
    public Integer getOriginalRow() {
        return row;
    }

    public String getDataColumn() {
        return dataColumn;
    }

    public CellAction getCellAction() {
        return action;
    }

    public static String replaceStuff(CellMapping mapping, String value) {
        String cellId = mapping.getColumn() + mapping.getOriginalRow();
        if (replaceMapping.containsKey(cellId)) {
            if (replaceMapping.get(cellId).containsKey(value)) {
                return replaceMapping.get(cellId).get(value);
            }
        }
        return "UNKNOWN_RESULT";
    }
}
