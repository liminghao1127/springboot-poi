/**
 * Copyright (C), 2019, 金科教育
 * FileName: BookServiceImpl
 * Author:   zyl
 * Date:     2019/8/8 14:25
 * History:
 * zyl          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */
package com.jk.service;

import com.jk.dao.BookDao;
import com.jk.model.Book;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * 〈一句话功能简述〉<br> 
 * 〈a〉
 *
 * @author zyl
 * @create 2019/8/8
 * @since 1.0.0
 */
@Service
public class BookServiceImpl implements BookService {

    @Autowired
    private BookDao bookDao;

    @Override
    public List<Book> findBookList(String bookname,String bookid,String bookprice,String booktime) {

        return bookDao.findBookList(bookname,bookid,bookprice,booktime);
    }

    @Override
    public List<Book> query() {
        return bookDao.query();
    }

    @Override
    public void addBook(List<Book> list) {
        bookDao.addBook( list);
    }
}
