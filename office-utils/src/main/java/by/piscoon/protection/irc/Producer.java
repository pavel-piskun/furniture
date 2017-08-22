/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package by.piscoon.protection.irc;


import java.io.*;

class Producer implements Runnable {

    private PrintWriter out;
    private Thread go;
    BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

    Producer(PrintWriter out) {
        this.out = out;

        go = new Thread(this);
        go.start();
    }

    public void run() {

        Thread th = Thread.currentThread();
        try {

            while (true) {
                out.println(br.readLine());
            }
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public void stop() {
        go = null;
    }
}