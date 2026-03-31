document.addEventListener('DOMContentLoaded', () => {
    const exportBtn = document.getElementById('exportBtn');
    if (!exportBtn) return;

    exportBtn.addEventListener('click', async () => {
        const overlay = document.getElementById('exportOverlay');
        const statusText = document.getElementById('exportStatus');
        overlay.classList.add('active');
        
        try {
            // 等待一下让 CSS 动画先渲染完毕
            await new Promise(r => setTimeout(r, 100));
            
            // 1. 初始化 pptx
            let pres = new PptxGenJS();
            pres.layout = 'LAYOUT_16x9'; // 设定 16:9 标准尺寸
            
            // 2. 获取所有幻灯片 DOM
            const slidesDom = document.querySelectorAll('.slide');
            const total = slidesDom.length;
            
            // 3. 逐一截图并当做背景图混入
            for(let i = 0; i < total; i++) {
                statusText.innerText = `正在对第 ${i+1}/${total} 页进行截图...`;
                const el = slidesDom[i];
                
                // 导出时仅固化必要的舞台背景，避免误改正文块的设计样式
                const overrideStyle = document.createElement('style');
                overrideStyle.innerHTML = `
                    .slide { background: #ffffff !important; background-color: #ffffff !important; box-shadow: none !important; border: none !important; }
                    .keyword, .outline-item.active { background-color: #f1f6fd !important; background: #f1f6fd !important; }
                `;
                document.head.appendChild(overrideStyle);

                const canvas = await html2canvas(el, {
                    scale: 1, 
                    useCORS: true,
                    logging: false
                    // 移除强制裁切坐标，由 html2canvas 自动根据无阴影的确切 bounds 裁切，防止偏移
                });
                
                // 截图完毕，无痕拔除强力覆盖样式
                document.head.removeChild(overrideStyle);

                const imgData = canvas.toDataURL('image/png');
                
                // 创建 PPTX 新页面并插入铺满画面的图片
                let slide = pres.addSlide();
                slide.addImage({ data: imgData, x: 0, y: 0, w: '100%', h: '100%' });
            }
            
            statusText.innerText = `正在打包生成 PPTX 文件...`;
            
            // 4. 保存为文件
            await pres.writeFile({ fileName: '智绘演示_开题答辩.pptx' });
            
            statusText.innerText = `✅ 导出成功！`;
            setTimeout(() => { overlay.classList.remove('active'); }, 1500);
        } catch (err) {
            console.error(err);
            statusText.innerText = `❌ 导出失败：请检查网络或控制台报错！`;
            setTimeout(() => { overlay.classList.remove('active'); }, 4000);
        }
    });
});
