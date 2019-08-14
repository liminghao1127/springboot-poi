/**
 * Copyright (C), 2015-2019, XXX有限公司
 * FileName: BookService
 * Author:   zyl
 * Date:     2019/8/8 14:25
 * History:
 */
package com.jk.service;


import com.jk.model.Book;

import java.util.List;

/**
 * 〈一句话功能简述〉<br> 
 * 〈〉
 *
 * @author zyl
 * @create 2019/8/8
 * @since 1.0.0
 */
public interface BookService {
    List<Book> query();

    void addBook(List<Book> list);

    List<Book> findBookList(String bookname,String bookid,String bookprice,String booktime);
}
