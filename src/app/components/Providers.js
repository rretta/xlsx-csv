'use client'

import { ChakraProvider } from '@chakra-ui/react'
import { CacheProvider } from '@chakra-ui/next-js'
import theme from '../theme'

export function Providers({ children }) {
  return (
    <CacheProvider>
      <ChakraProvider theme={theme}>
        {children}
      </ChakraProvider>
    </CacheProvider>
  )
} 