package com.xt.utils;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

public class ThreadHelper {
    public static ThreadInfo getTreadInfo(Thread thread) {
        return ThreadInfo.builder()
                .id(thread.getId())
                .name(thread.getName())
                .status(thread.getState().toString())
                .priority(thread.getPriority())
                .isDaemon(thread.isDaemon())
                .build();
    }

    public static List<ThreadInfo> getThreadInfoList() {
        List<ThreadInfo> threadInfoList = new ArrayList<>();
        Set<Thread> threadSet = Thread.getAllStackTraces().keySet();
        for (Thread t : threadSet) {
            threadInfoList.add(getTreadInfo(t));
        }
        return threadInfoList;
    }

    public static void printAllThreads() {
        Set<Thread> threadSet = Thread.getAllStackTraces().keySet();
        System.out.printf("%-34s \t %-15s \t %-15s \t %s\n", "Name", "State", "Priority", "isDaemon");
        // iterating over the threads to get the names of all the active threads
        for (Thread t : threadSet) {
            System.out.printf("%-34s \t %-15s \t %-15d \t %s\n", t.getName(), t.getState(), t.getPriority(), t.isDaemon());
        }
    }
}
