# 测试用的二进制样本

| 文件 | 来源 | 用途 |
| --- | --- | --- |
| `agile-msoffcrypto.docx` | 用 docx-view 生成一份 .docx，再用 [msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool) 6.0 以密码 `参考 Reference-1` 加密（Agile，AES-256 / SHA-512 / 100000 次） | 用一份**别的实现**产出的加密文件验证解密，不让测试只验证「自己加密、自己解密」。正文超过 4096 字节，`EncryptedPackage` 落在普通扇区里——msoffcrypto-tool 6.0 写小于 4096 字节的流时会把迷你流的内容放错，连它自己都解不开那样的文件。 |
