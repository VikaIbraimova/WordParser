package ru.tempWorld5;

//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;
import ru.tempWorld5.exceptions.ParseException;

import javax.swing.*;

public class Main {

   // private static Logger log = LoggerFactory.getLogger(Main.class);
    private static ApplicationContext applicationContext = new ClassPathXmlApplicationContext("spring-config.xml");
   // private static Parser parser = applicationContext.getBean(Parser.class);
    //Пишет в Excel, правда пока без учета адресов
    //private static Parser5 parser = applicationContext.getBean(Parser5.class);
    //Добавляет еще в Excel строку, в туже колонку, куда записал данные Parser5
    //private static Parser6 parser = applicationContext.getBean(Parser6.class);

    //Пишет в нужные ячейки, правда пока строки скачут
    private static Parser7 parser = applicationContext.getBean(Parser7.class);

    //Выводит инфрмацию, разобрав Word
    //private static Parser8 parser = applicationContext.getBean(Parser8.class);

    //Пока ерунда
    //private static Parser9 parser = applicationContext.getBean(Parser9.class);

    //Немного изменив алгоритм,тепрь добавляется только один элмент
    //private static Parser10 parser = applicationContext.getBean(Parser10.class);

    public static void main(String[] args) {
        try {
            //Берем документ Word, парсим его
            //parser.parse();
            //Выводим распарсенные данные в Excel
            //parser.parseOutExcel(parser.parse());
            //Пишем в Excel
            parser.parseOutExcel2(parser.parse());
        } catch (ParseException exception) {
            String message = exception.getMessage();
            if (exception.getCause() != null) {
                message = message + exception.getCause();
            }
            JOptionPane.showMessageDialog(null, "Parse exception:\n" + message);
        }
    }
}