package org.wymix.excel.tools;

import lombok.Data;
import lombok.ToString;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

@Data
@ToString
public class ExcelInfo {

    private String path;

    private Sheet sheet;

    @Data
    public static class Sheet {
        private Integer index = 0;

        private Title title;
    }

    @Data
    public static class Title {
        private Integer index = 0;

        private List<Column> column;
    }

    @Data
    public static class Column {
        private String value;
        private String property;
        private Integer index;
        private boolean nullable;
        private boolean unique;
        private String dataFormat;
        private String match;

        private Map<String, Integer> uniqueVals = new HashMap<>();

        private Map<Integer, String> errorMsg = new HashMap<>();

        final void validate(int rowIndex, Object cellVal) {
            String str = cellVal == null ? null : String.valueOf(cellVal);
            StringBuffer errorMsg = new StringBuffer();
            if (!this.nullable && (str == null || str.length() == 0)) {
                errorMsg.append("不能为空; \n");
            }

            if (this.unique && uniqueVals.putIfAbsent(str, rowIndex) != null) {
                errorMsg.append("不能重复; \n");
            }

            if (null!=this.match&&!Pattern.compile(this.match).matcher(str).matches()) {
                errorMsg.append("格式不匹配; \n");
            }
            if(errorMsg.length()>0){
                this.errorMsg.put(rowIndex, errorMsg.toString());
            }
        }

    }

}
