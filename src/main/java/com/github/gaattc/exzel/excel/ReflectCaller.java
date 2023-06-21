package com.github.gaattc.exzel.excel;

import org.joor.Reflect;

import java.util.regex.Pattern;

/**
 * @author gaattc
 * @version 6.0
 * Created by gaattc on 2023/6/20
 */
public class ReflectCaller {

    private static final String SPLITTER = "#";
    private static final String EMPTY = "";
    private static final Pattern PATTERN = Pattern.compile("((\\S+)#(\\S+))");

    /**
     * {@link ExcelMapping#contentConverter()}
     */
    public static Object convert(String methodFullName, Object origin) {
        try {
            if (!checkMethodCorrect(methodFullName)) {
                return origin;
            }
            String[] split = methodFullName.split(SPLITTER);
            return Reflect.onClass(split[0])
                    .call(split[1], origin)
                    .get();
        } catch (Throwable ignore) {
            return origin;
        }
    }

    /**
     * {@link ExcelStyle#columnNameSupplier()}
     */
    public static String supplier(String methodFullName) {
        try {
            if (!checkMethodCorrect(methodFullName)) {
                return EMPTY;
            }
            String[] split = methodFullName.split(SPLITTER);
            return Reflect.onClass(split[0])
                    .call(split[1])
                    .get();
        } catch (Throwable ignore) {
            return EMPTY;
        }
    }

    private static boolean checkMethodCorrect(String methodFullName) {
        return PATTERN.matcher(methodFullName).matches();
    }

}
