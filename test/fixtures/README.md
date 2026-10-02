# 测试用的二进制样本

| 文件 | 来源 | 用途 |
| --- | --- | --- |
| `agile-msoffcrypto.docx` | 用 docx-view 生成一份 .docx，再用 [msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool) 6.0 以密码 `参考 Reference-1` 加密（Agile，AES-256 / SHA-512 / 100000 次） | 用一份**别的实现**产出的加密文件验证解密，不让测试只验证「自己加密、自己解密」。正文超过 4096 字节，`EncryptedPackage` 落在普通扇区里——msoffcrypto-tool 6.0 写小于 4096 字节的流时会把迷你流的内容放错，连它自己都解不开那样的文件。 |
| `standard-office2007.docx` | 取自 [msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool) 的测试样本 `tests/inputs/ecma376standard_password.docx`（MIT 许可，© 2015 nolze），密码 `Password1234_`。EncryptionInfo 里的 CSP 是「Microsoft Enhanced RSA and AES Cryptographic Provider」，是 Office 产出的真实文件 | 验证 Office 2007「Standard」加密（AES-128 + SHA-1）的解密。按第 26 条先核对过参照：msoffcrypto-tool 解出来的结果与它仓库里的已知明文逐字节一致，我们解出来的也与那份明文逐字节一致。 |
