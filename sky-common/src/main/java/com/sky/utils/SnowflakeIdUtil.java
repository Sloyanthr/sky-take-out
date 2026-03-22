package com.sky.utils;

/**
 * 雪花算法工具类
 * 用于生成全局唯一的订单号、ID等
 */
public class SnowflakeIdUtil {

    /**
     * 起始时间戳 (2024-01-01 00:00:00)
     * 一旦确定请勿修改，否则可能导致 ID 重复
     */
    private static final long TWEPOCH = 1704067200000L;

    /**
     * 机器标识占用的位数
     */
    private static final long WORKER_ID_BITS = 5L;
    private static final long DATACENTER_ID_BITS = 5L;

    /**
     * 序列号占用的位数
     */
    private static final long SEQUENCE_BITS = 12L;

    /**
     * 毫秒内序列的最大值 (4095)
     */
    private static final long SEQUENCE_MASK = -1L ^ (-1L << SEQUENCE_BITS);

    private static final long WORKER_ID_SHIFT = SEQUENCE_BITS;
    private static final long DATACENTER_ID_SHIFT = SEQUENCE_BITS + WORKER_ID_BITS;
    private static final long TIMESTAMP_LEFT_SHIFT = SEQUENCE_BITS + WORKER_ID_BITS + DATACENTER_ID_BITS;

    private static long workerId = 1L;      // 终端ID
    private static long datacenterId = 1L;  // 数据中心ID
    private static long sequence = 0L;      // 毫秒内序列
    private static long lastTimestamp = -1L; // 上次生成ID的时间截

    /**
     * 静态内部实例（单例模式）
     */
    private static final SnowflakeIdUtil instance = new SnowflakeIdUtil();

    private SnowflakeIdUtil() {}

    /**
     * 生成下一个 ID (Long类型)
     */
    public static synchronized long nextId() {
        long timestamp = timeGen();

        // 如果当前时间小于上一次ID生成的时间戳，说明系统时钟回退过，抛出异常
        if (timestamp < lastTimestamp) {
            throw new RuntimeException(
                    String.format("Clock moved backwards. Refusing to generate id for %d milliseconds", lastTimestamp - timestamp));
        }

        // 如果是同一时间生成的，则进行毫秒内序列
        if (lastTimestamp == timestamp) {
            sequence = (sequence + 1) & SEQUENCE_MASK;
            // 毫秒内序列溢出
            if (sequence == 0) {
                // 阻塞到下一个毫秒,获得新的时间戳
                timestamp = tilNextMillis(lastTimestamp);
            }
        } else {
            // 时间戳改变，毫秒内序列重置
            sequence = 0L;
        }

        lastTimestamp = timestamp;

        // 移位并通过或运算拼到一起组成64位的ID
        return ((timestamp - TWEPOCH) << TIMESTAMP_LEFT_SHIFT)
                | (datacenterId << DATACENTER_ID_SHIFT)
                | (workerId << WORKER_ID_SHIFT)
                | sequence;
    }

    /**
     * 生成下一个 ID 的字符串形式 (最推荐用于订单号)
     */
    public static String nextIdStr() {
        return String.valueOf(nextId());
    }

    /**
     * 阻塞到下一个毫秒，直到获得新的时间戳
     */
    private static long tilNextMillis(long lastTimestamp) {
        long timestamp = timeGen();
        while (timestamp <= lastTimestamp) {
            timestamp = timeGen();
        }
        return timestamp;
    }

    /**
     * 返回以毫秒为单位的当前时间
     */
    private static long timeGen() {
        return System.currentTimeMillis();
    }
}