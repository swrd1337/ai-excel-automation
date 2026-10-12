import * as TooltipPrimitive from '@radix-ui/react-tooltip';
import type { ComponentProps, ReactNode } from 'react';
import { cn } from '@/lib/utils';

const TooltipProvider = TooltipPrimitive.Provider;
const Tooltip = TooltipPrimitive.Root;
const TooltipTrigger = TooltipPrimitive.Trigger;

function TooltipContent({ className, sideOffset = 6, ...props }: ComponentProps<typeof TooltipPrimitive.Content>) {
  return (
    <TooltipPrimitive.Portal>
      <TooltipPrimitive.Content
        sideOffset={sideOffset}
        className={cn(
          'z-50 max-w-[19rem] rounded-md border bg-popover px-2.5 py-1.5 text-xs leading-relaxed text-popover-foreground shadow-md',
          className
        )}
        {...props}
      />
    </TooltipPrimitive.Portal>
  );
}

/**
 * Convenience wrapper: most call sites are just (label, text).
 *
 * `asChild` is not optional here — TooltipTrigger otherwise renders its own
 * <button>, which would nest a second button around every control it wraps
 * (invalid markup, and focus lands on the wrapper instead of the control).
 */
export function Hint({
  children,
  text,
  side = 'bottom',
}: {
  children: ReactNode;
  text: ReactNode;
  side?: ComponentProps<typeof TooltipContent>['side'];
}) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <span className="inline-flex">{children}</span>
      </TooltipTrigger>
      <TooltipContent side={side}>{text}</TooltipContent>
    </Tooltip>
  );
}

export { Tooltip, TooltipTrigger, TooltipContent, TooltipProvider };
