package com.github.gaattc.exzel.excel;

import java.lang.reflect.Method;
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
            Class<?> clazz = Class.forName(split[0]);
            Method method = clazz.getMethod(split[1], origin.getClass());
            return method.invoke(null, origin);
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
            Class<?> clazz = Class.forName(split[0]);
            Method method = clazz.getMethod(split[1]);
            return ((String) method.invoke(null));
        } catch (Throwable ignore) {
            return EMPTY;
        }
    }

    private static boolean checkMethodCorrect(String methodFullName) {
        return PATTERN.matcher(methodFullName).matches();
    }

}
