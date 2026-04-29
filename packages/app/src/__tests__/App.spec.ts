import { describe, it, expect } from 'vitest'
import { mount } from '@vue/test-utils'

import App from '../App.vue'

describe('App', () => {
  it('renders the Phase 0 parser UI', () => {
    const wrapper = mount(App)

    expect(wrapper.text()).toContain('企业可控 PPTX Web Player')
    expect(wrapper.text()).toContain('上传本地 PPTX')
    expect(wrapper.text()).toContain('Diagnostics')
    expect(wrapper.text()).toContain('区级平台介绍')
    expect(wrapper.text()).toContain('Slidesgo AI Tech')
    expect(wrapper.text()).toContain('请选择一个样本或上传 PPTX 文件开始解析')
  })
})
