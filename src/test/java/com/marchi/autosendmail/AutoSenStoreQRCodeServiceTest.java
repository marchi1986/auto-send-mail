package com.marchi.autosendmail;

import com.marchi.autosendmail.service.AutoSendStoreQRCodeService;
import org.junit.jupiter.api.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

@SpringBootTest
public class AutoSenStoreQRCodeServiceTest<AutoSenStoreQRCodeService> {

    @Autowired
    private AutoSendStoreQRCodeService autoSenStoreQRCodeService;

    @Test
    public void test(){
        autoSenStoreQRCodeService.read();
    }
}
