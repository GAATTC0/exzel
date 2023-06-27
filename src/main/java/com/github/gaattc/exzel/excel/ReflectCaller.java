package com.github.gaattc.exzel.excel;

import lombok.extern.slf4j.Slf4j;
import org.joor.Reflect;

import java.util.regex.Pattern;

/**
 * @author gaattc
 * @since 1.0
 * Created by gaattc on 2023/6/20
 */
@Slf4j
public class ReflectCaller {

    private static final String SPLITTER = "#";
    private static final String EMPTY = "";
    private static final Pattern PATTERN = Pattern.compile("((\\S+)#(\\S+))");

    /**
     * {@link ExcelMapping#contentConverter()}
     */
    public static Object function(String methodFullName, Object origin, ClassLoader classLoader) {
        try {
            if (!checkMethodCorrect(methodFullName)) {
                return origin;
            }
            String[] split = methodFullName.split(SPLITTER);
            return Reflect.onClass(split[0], classLoader)
                    .call(split[1], origin)
                    .get();
        } catch (Throwable ignore) {
            return origin;
        }
    }

    /**
     * {@link ExcelStyle#columnNameSupplier()}
     */
    public static String supplier(String methodFullName, ClassLoader classLoader) {
        try {
            if (!checkMethodCorrect(methodFullName)) {
                return EMPTY;
            }
            String[] split = methodFullName.split(SPLITTER);
            return Reflect.onClass(split[0], classLoader)
                    .call(split[1])
                    .get();
        } catch (Throwable ignore) {
            return EMPTY;
        }
    }

    @SuppressWarnings("BooleanMethodIsAlwaysInverted")
    private static boolean checkMethodCorrect(String methodFullName) {
        boolean matches = PATTERN.matcher(methodFullName).matches();
        if (!matches) {
            log.warn("incorrect full method name syntax: {}, ignored", methodFullName);
        }
        return matches;
    }

}
